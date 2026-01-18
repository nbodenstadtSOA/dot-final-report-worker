import JSZip from "jszip";

const TEMPLATE_KEY = "templates/final-report-template.xlsm";

// =======================
// FIELD ORDERS (EXACT)
// =======================

const SCENARIO_FIELDS = [
  "Name",
  "Appr Type",
  "Status",
  "Component",
  "Appropriation Type (from Component)",
  "1000 Expenditures",
  "2000 Expenditures",
  "3000 Expenditures",
  "4000 Expenditures",
  "5000 Expenditures",
  "Final Notes",
  "Fiscal Year 2",
  "Month 2",
  "Actuals Date (from Month 2)",
  "Personal Services Projection Description",
  "Projection Mode",
  "Created By",
  "Created Time",
];

const LINES_FIELDS = [
  "Name",
  "Projection Type",
  "Object Class (from Object Class)",
  "Obj. Type (from Object Class)",
  "Obj. Group (from Object Class)",
  "Object Class with Name",
  "Personal Services?",
  "Pre-Encumbrance",
  "Encumbrance",
  "Expenditure",
  "Expected Expenditures",
  "Total Expenditures",
  "Total Plan (Manual)",
  "Expected Expenditures (Calc)",
  "Notes",
  "RSA Budget",
  "Program Code",
  "RSA Description",

  // NEW (Column S)
  "PY Actuals",
];

const SUBLINES_FIELDS = [
  "Name",
  "Projection Lines",
  "Object Class",
  "Pre-Encumbrances",
  "Encumbrances",
  "Expenditures",
  "Projected Expenditures",
  "Total Projected Spend",
  "Notes",
  "District Note",
  "Total Plan (Manual)",
  "Projected Expenditures (Calc)",
  "Total Expenditures (Calc)",
];

const FUNDSOURCES_FIELDS = [
  "Appr Unit",
  "Fund",
  "Expected Revenue",
  "1000",
  "2000",
  "3000",
  "4000",
  "5000",
  "Total Expenditures",
  "Balance",
  "1000 Exp Budget",
  "2000 Exp Budget",
  "3000 Exp Budget",
  "4000 Exp Budget",
  "5000 Exp Budget",
  "1000 Pending Budget Changes",
  "2000 Pending Budget Changes",
  "3000 Pending Budget Changes",
  "4000 Pending Budget Changes",
  "5000 Pending Budget Changes",
  "1000 Balance",
  "2000 Balance",
  "3000 Balance",
  "4000 Balance",
  "5000 Balance",
  "Support Lines Total Budget",
  "Support Lines Balance",
  "Support Lines Expenditures",
  "Expenditure Budget",
  "Budget Change Notes",
  "Balance (Exp Budget)",
];

// =======================
// NUMERIC COLUMN SETS
// (write as numbers, not text)
// =======================

const NUMERIC_FIELDS_BY_SHEET = {
  "Data_Scenario": new Set([
    "1000 Expenditures",
    "2000 Expenditures",
    "3000 Expenditures",
    "4000 Expenditures",
    "5000 Expenditures",
  ]),

  "Data_Lines": new Set([
    "Pre-Encumbrance",
    "Encumbrance",
    "Expenditure",
    "Expected Expenditures",
    "Total Expenditures",
    "Total Plan (Manual)",
    "Expected Expenditures (Calc)",
    "RSA Budget",
    "PY Actuals", // NEW
  ]),

  "Data_Sub_Lines": new Set([
    "Pre-Encumbrances",
    "Encumbrances",
    "Expenditures",
    "Projected Expenditures",
    "Total Projected Spend",
    "Total Plan (Manual)",
    "Projected Expenditures (Calc)",
    "Total Expenditures (Calc)",
  ]),

  "Data_Fund-Sources": new Set([
    "Expected Revenue",
    "1000", "2000", "3000", "4000", "5000",
    "Total Expenditures",
    "Balance",
    "1000 Exp Budget", "2000 Exp Budget", "3000 Exp Budget", "4000 Exp Budget", "5000 Exp Budget",
    "1000 Pending Budget Changes", "2000 Pending Budget Changes", "3000 Pending Budget Changes", "4000 Pending Budget Changes", "5000 Pending Budget Changes",
    "1000 Balance", "2000 Balance", "3000 Balance", "4000 Balance", "5000 Balance",
    "Support Lines Total Budget",
    "Support Lines Balance",
    "Support Lines Expenditures",
    "Expenditure Budget",
    "Balance (Exp Budget)",
  ]),
};

// =======================
// WORKER
// =======================

export default {
  async fetch(request, env) {
    // Hard fail only for non-POST
    if (request.method !== "POST") {
      return new Response("Method Not Allowed", { status: 405 });
    }

    try {
      // --- Auth ---
      const apiKeyHeader = request.headers.get("x-api-key");
      if (!apiKeyHeader) return new Response("Forbidden: missing x-api-key", { status: 403 });
      if (!env.API_KEY) return new Response("Forbidden: env.API_KEY not configured", { status: 500 });
      if (apiKeyHeader !== env.API_KEY) return new Response("Forbidden: api key mismatch", { status: 403 });

      // --- Required bindings ---
      if (!env.TEMPLATES_BUCKET) return jsonError(500, "Missing binding: TEMPLATES_BUCKET");
      if (!env.REPORTS_BUCKET) return jsonError(500, "Missing binding: REPORTS_BUCKET");
      if (!env.R2_PUBLIC_BASE) return jsonError(500, "Missing env var: R2_PUBLIC_BASE");

      // --- Parse JSON body ---
      let payload;
      try {
        payload = await request.json();
      } catch {
        return jsonError(400, "Invalid JSON body");
      }

      // Accept payload wrapped as: { _payloadJson: "..." }
      if (typeof payload?._payloadJson === "string") {
        const parsed = tryParseJson(payload._payloadJson);
        if (parsed && typeof parsed === "object") payload = parsed;
      }

      const scenarioId = payload?.scenarioId;
      const scenarioName = payload?.scenarioName ?? "Final_Report";

      if (!scenarioId) return jsonError(400, "Missing scenarioId");

      // Normalize everything (Make-proof)
      const scenario = normalizeScenario(payload?.scenario);
      const projectionLines = normalizeRows(payload?.projectionLines);
      const subLines = normalizeRows(payload?.subLines);
      const fundSources = normalizeRows(payload?.fundSources);

      console.log("scenarioId:", scenarioId);
      console.log("scenario keys:", Object.keys(scenario || {}).length);
      console.log("projectionLines:", projectionLines.length, "firstType:", typeof projectionLines[0]);
      console.log("subLines:", subLines.length, "firstType:", typeof subLines[0]);
      console.log("fundSources:", fundSources.length, "firstType:", typeof fundSources[0]);

      // --- Load template from R2 ---
      const templateObj = await env.TEMPLATES_BUCKET.get(TEMPLATE_KEY);
      if (!templateObj) return jsonError(500, `Template not found in TEMPLATES_BUCKET: ${TEMPLATE_KEY}`);

      const templateBytes = new Uint8Array(await templateObj.arrayBuffer());
      const zip = await JSZip.loadAsync(templateBytes);

      // Preload workbook mappings
      const workbookXml = await readText(zip, "xl/workbook.xml");
      const workbookRelsXml = await readText(zip, "xl/_rels/workbook.xml.rels");

      // --- Write all 4 tables ---
      await writeTable({
        zip,
        workbookXml,
        workbookRelsXml,
        sheetName: "Data_Scenario",
        fields: SCENARIO_FIELDS,
        rows: [scenario], // always one row
      });

      await writeTable({
        zip,
        workbookXml,
        workbookRelsXml,
        sheetName: "Data_Lines",
        fields: LINES_FIELDS,
        rows: projectionLines,
      });

      await writeTable({
        zip,
        workbookXml,
        workbookRelsXml,
        sheetName: "Data_Sub_Lines",
        fields: SUBLINES_FIELDS,
        rows: subLines,
      });

      await writeTable({
        zip,
        workbookXml,
        workbookRelsXml,
        sheetName: "Data_Fund-Sources",
        fields: FUNDSOURCES_FIELDS,
        rows: fundSources,
      });

      // Re-zip (macros preserved)
      const outBytes = await zip.generateAsync({ type: "uint8array" });

      // --- Save output to REPORTS_BUCKET ---
      const safeScenario = String(scenarioName).replace(/[^\w\-]+/g, "_").slice(0, 120);
      const ts = new Date().toISOString().replace(/[:.]/g, "-");
      const fileName = `${ts}-${safeScenario}.xlsm`;
      const outKey = `reports/${scenarioId}/${fileName}`;

      await env.REPORTS_BUCKET.put(outKey, outBytes, {
        httpMetadata: { contentType: "application/vnd.ms-excel.sheet.macroEnabled.12" },
      });

      const fileUrl = `${env.R2_PUBLIC_BASE}/${outKey.split("/").map(encodeURIComponent).join("/")}`;

      return new Response(JSON.stringify({ fileUrl, fileName }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    } catch (err) {
      // Always return JSON, never HTML 1101 to Make
      console.log("Unhandled exception:", err?.stack || String(err));
      return jsonError(500, `Worker exception: ${err?.message || String(err)}`);
    }
  },
};

// =======================
// TABLE WRITER
// =======================

async function writeTable({ zip, workbookXml, workbookRelsXml, sheetName, fields, rows }) {
  const numericSet = NUMERIC_FIELDS_BY_SHEET[sheetName] || new Set();

  // Resolve sheet XML path
  const sheetRid = findSheetRidByName(workbookXml, sheetName);
  const sheetPath = resolveWorkbookRidToTarget(workbookRelsXml, sheetRid); // e.g. worksheets/sheet3.xml
  const fullSheetPath = "xl/" + sheetPath;

  let sheetXml = await readText(zip, fullSheetPath);

  // Resolve table XML path (first table relationship on the sheet)
  const sheetRelsPath = `xl/worksheets/_rels/${basename(sheetPath)}.rels`;
  const sheetRelsXml = await readText(zip, sheetRelsPath);

  const tableTarget = findFirstTableTarget(sheetRelsXml); // ../tables/tableX.xml
  const tablePath = normalizePath("xl/worksheets", tableTarget); // xl/tables/tableX.xml
  let tableXml = await readText(zip, tablePath);

  // Build matrix in field order, converting numeric columns to real numbers when possible
  const matrix = rows.map((obj) =>
    fields.map((f) => {
      const raw = obj?.[f] ?? "";
      if (numericSet.has(f)) {
        const num = parseMoneyNumber(raw);
        return num == null ? "" : num;
      }
      return String(raw ?? "");
    })
  );

  // Replace all rows starting at A2
  sheetXml = replaceRowsTyped(sheetXml, 2, matrix, fields, sheetName);

  // Update dimension + table ref to include header + N data rows (min 1)
  const lastCol = colLetter(fields.length - 1);
  const dataRows = Math.max(1, matrix.length);
  const lastRow = 1 + dataRows; // header row 1 + data rows
  const ref = `A1:${lastCol}${lastRow}`;

  sheetXml = upsertDimension(sheetXml, ref);
  tableXml = updateTableRef(tableXml, ref);

  // Save back into ZIP
  zip.file(fullSheetPath, sheetXml);
  zip.file(tablePath, tableXml);
}

// =======================
// TYPE HANDLING
// =======================

function parseMoneyNumber(value) {
  if (value == null) return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const s0 = String(value).trim();
  if (s0 === "") return null;

  // Allow negatives written as (123.45)
  const isParenNeg = /^\(.*\)$/.test(s0);

  // Remove parens, $ and commas, and whitespace
  const cleaned = s0
    .replace(/^\(/, "")
    .replace(/\)$/, "")
    .replace(/\$/g, "")
    .replace(/,/g, "")
    .replace(/\s+/g, "");

  const n = Number(cleaned);
  if (!Number.isFinite(n)) return null;
  return isParenNeg ? -n : n;
}

// =======================
// MAKE-PROOF NORMALIZATION
// =======================

function tryParseJson(value) {
  if (typeof value !== "string") return null;
  const t = value.trim();
  if (!(t.startsWith("{") || t.startsWith("["))) return null;
  try {
    return JSON.parse(t);
  } catch {
    return null;
  }
}

function normalizeScenario(value) {
  let v = value;

  if (Array.isArray(v) && v.length === 1) v = v[0];

  if (typeof v === "string") {
    const parsed = tryParseJson(v);
    if (parsed && typeof parsed === "object") v = parsed;
  }

  if (Array.isArray(v) && v.length === 1 && typeof v[0] === "object") v = v[0];

  if (v && typeof v === "object" && !Array.isArray(v)) return v;

  // Never hard fail: blank object still lets workbook build
  return {};
}

function normalizeRows(value) {
  const parsed = tryParseJson(value);
  if (parsed && Array.isArray(parsed)) return parsed;

  if (Array.isArray(value)) {
    // Sometimes Make sends ["[{...},{...}]"]
    if (value.length === 1 && typeof value[0] === "string") {
      const inner = tryParseJson(value[0]);
      if (inner && Array.isArray(inner)) return inner;
    }
    // Or items are JSON strings
    return value
      .map((item) => tryParseJson(item) ?? item)
      .filter((x) => x && typeof x === "object" && !Array.isArray(x));
  }

  return [];
}

// =======================
// XML WRITING (typed)
// =======================

function replaceRowsTyped(sheetXml, startRow, valuesMatrix, fields, sheetName) {
  const start = Number(startRow);
  const numericSet = NUMERIC_FIELDS_BY_SHEET[sheetName] || new Set();

  return sheetXml.replace(/<sheetData>([\s\S]*?)<\/sheetData>/i, (m, inner) => {
    // Remove existing rows with r >= startRow
    const cleaned = inner.replace(/<row\b[^>]*\br="(\d+)"[\s\S]*?<\/row>/gi, (rowXml, rStr) => {
      const r = Number(rStr);
      return r >= start ? "" : rowXml;
    });

    // Create new rows
    const rowsXml = valuesMatrix.map((vals, i) => {
      const r = start + i;

      const cells = vals
        .map((v, idx) => {
          const fieldName = fields[idx];
          const addr = colLetter(idx) + r;

          if (numericSet.has(fieldName) && typeof v === "number") {
            return makeNumberCell(addr, v);
          }
          return makeInlineStrCell(addr, v == null ? "" : String(v));
        })
        .join("");

      return `<row r="${r}">${cells}</row>`;
    });

    // Always keep at least 1 data row so the table isn't zero-length
    if (rowsXml.length === 0) {
      rowsXml.push(`<row r="${start}"></row>`);
    }

    return `<sheetData>${cleaned.trimEnd()}${rowsXml.join("")}</sheetData>`;
  });
}

function makeInlineStrCell(r, text) {
  const t = escapeXml(text ?? "");
  return `<c r="${r}" t="inlineStr"><is><t>${t}</t></is></c>`;
}

function makeNumberCell(r, n) {
  // Numeric cell: no t="inlineStr"
  return `<c r="${r}"><v>${n}</v></c>`;
}

// =======================
// ZIP / XML UTILITIES
// =======================

function jsonError(status, message) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

async function readText(zip, path) {
  const f = zip.file(path);
  if (!f) throw new Error(`Missing zip entry: ${path}`);
  return await f.async("text");
}

function basename(p) {
  return p.split("/").pop();
}

function normalizePath(fromDir, relTarget) {
  const fromParts = fromDir.split("/").filter(Boolean);
  const relParts = relTarget.split("/").filter(Boolean);
  const stack = [...fromParts];
  for (const part of relParts) {
    if (part === "..") stack.pop();
    else stack.push(part);
  }
  return stack.join("/");
}

function findSheetRidByName(workbookXml, sheetName) {
  const re = new RegExp(`<sheet[^>]*name="${escapeReg(sheetName)}"[^>]*r:id="([^"]+)"`, "i");
  const m = workbookXml.match(re);
  if (!m) throw new Error(`Sheet not found in workbook.xml: ${sheetName}`);
  return m[1];
}

function resolveWorkbookRidToTarget(workbookRelsXml, rid) {
  const re = new RegExp(`<Relationship[^>]*Id="${escapeReg(rid)}"[^>]*Target="([^"]+)"`, "i");
  const m = workbookRelsXml.match(re);
  if (!m) throw new Error(`workbook.xml.rels missing target for ${rid}`);
  return m[1];
}

function findFirstTableTarget(sheetRelsXml) {
  const re = /<Relationship[^>]*Type="[^"]*\/table"[^>]*Target="([^"]+)"/i;
  const m = sheetRelsXml.match(re);
  if (!m) throw new Error("No table relationship found for this sheet");
  return m[1];
}

function upsertDimension(sheetXml, ref) {
  if (sheetXml.includes("<dimension")) {
    return sheetXml.replace(/<dimension[^>]*ref="[^"]*"[^\/]*\/>/i, `<dimension ref="${ref}"/>`);
  }
  return sheetXml.replace(/<worksheet[^>]*>/i, (m) => `${m}\n<dimension ref="${ref}"/>`);
}

function updateTableRef(tableXml, ref) {
  tableXml = tableXml.replace(/<table([^>]*?)ref="[^"]*"/i, `<table$1ref="${ref}"`);
  if (tableXml.includes("<autoFilter")) {
    tableXml = tableXml.replace(/<autoFilter([^>]*?)ref="[^"]*"/i, `<autoFilter$1ref="${ref}"`);
  }
  return tableXml;
}

function colLetter(idx) {
  let n = idx + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeReg(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
