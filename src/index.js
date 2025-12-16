import JSZip from "jszip";

const TEMPLATE_KEY = "templates/final-report-template.xlsm";

// ---------- Column orders (MUST match Excel table header order) ----------

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

// ---------- Worker ----------

export default {
  async fetch(request, env) {
    if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });

    // Auth
    const apiKeyHeader = request.headers.get("x-api-key");
    if (!apiKeyHeader) return new Response("Forbidden: missing x-api-key", { status: 403 });
    if (!env.API_KEY) return new Response("Forbidden: env.API_KEY not configured", { status: 500 });
    if (apiKeyHeader !== env.API_KEY) return new Response("Forbidden: api key mismatch", { status: 403 });

    // Parse payload
    let payload;
    try {
      payload = await request.json();
    } catch {
      return new Response("Invalid JSON", { status: 400 });
    }

    // Accept either:
    // A) direct payload {scenarioId, scenarioName, scenario, projectionLines, ...}
    // B) payload wrapped as { _payloadJson: "...." } (Make-proof)
    if (typeof payload?._payloadJson === "string") {
      const parsed = tryParseJson(payload._payloadJson);
      if (parsed && typeof parsed === "object") payload = parsed;
    }

    const scenarioId = payload?.scenarioId;
    const scenarioName = payload?.scenarioName;

    if (!scenarioId) return new Response("Missing scenarioId", { status: 400 });

    // Make-proof normalization
    const scenario = normalizeScenario(payload?.scenario);
    const projectionLines = normalizeRows(payload?.projectionLines);
    const subLines = normalizeRows(payload?.subLines);
    const fundSources = normalizeRows(payload?.fundSources);

    // Debug signals (leave on until stable)
    console.log("scenarioId:", scenarioId);
    console.log("scenario type:", typeof payload?.scenario, "isArray:", Array.isArray(payload?.scenario));
    console.log("projectionLines:", Array.isArray(projectionLines), projectionLines.length, "firstType:", typeof projectionLines[0]);
    console.log("subLines:", Array.isArray(subLines), subLines.length, "firstType:", typeof subLines[0]);
    console.log("fundSources:", Array.isArray(fundSources), fundSources.length, "firstType:", typeof fundSources[0]);

    // Load template
    const templateObj = await env.TEMPLATES_BUCKET.get(TEMPLATE_KEY);
    if (!templateObj) return new Response(`Template not found: ${TEMPLATE_KEY}`, { status: 500 });

    const templateBytes = new Uint8Array(await templateObj.arrayBuffer());
    const zip = await JSZip.loadAsync(templateBytes);

    // Resolve workbook sheet relationships once
    const workbookXml = await readText(zip, "xl/workbook.xml");
    const workbookRelsXml = await readText(zip, "xl/_rels/workbook.xml.rels");

    // Write each table
    await writeTable({
      zip,
      workbookXml,
      workbookRelsXml,
      sheetName: "Data_Scenario",
      fields: SCENARIO_FIELDS,
      rows: [scenario], // always 1 row (blank object is OK)
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

    // Re-zip
    const outBytes = await zip.generateAsync({ type: "uint8array" });

    // Save output
    const safeScenario = (scenarioName || "Final_Report")
      .toString()
      .replace(/[^\w\-]+/g, "_")
      .slice(0, 120);

    const ts = new Date().toISOString().replace(/[:.]/g, "-");
    const fileName = `${ts}-${safeScenario}.xlsm`;
    const outKey = `reports/${scenarioId}/${fileName}`;

    await env.REPORTS_BUCKET.put(outKey, outBytes, {
      httpMetadata: { contentType: "application/vnd.ms-excel.sheet.macroEnabled.12" },
    });

    if (!env.R2_PUBLIC_BASE) return new Response("Missing env.R2_PUBLIC_BASE", { status: 500 });
    const fileUrl = `${env.R2_PUBLIC_BASE}/${outKey.split("/").map(encodeURIComponent).join("/")}`;

    return new Response(JSON.stringify({ fileUrl, fileName }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  },
};

// ---------- Core write helper ----------

async function writeTable({ zip, workbookXml, workbookRelsXml, sheetName, fields, rows }) {
  const sheetRid = findSheetRidByName(workbookXml, sheetName);
  const sheetPath = resolveWorkbookRidToTarget(workbookRelsXml, sheetRid);
  const fullSheetPath = "xl/" + sheetPath;

  let sheetXml = await readText(zip, fullSheetPath);

  const sheetRelsPath = `xl/worksheets/_rels/${basename(sheetPath)}.rels`;
  const sheetRelsXml = await readText(zip, sheetRelsPath);

  const tableTarget = findFirstTableTarget(sheetRelsXml);
  const tablePath = normalizePath("xl/worksheets", tableTarget);
  let tableXml = await readText(zip, tablePath);

  const matrix = rows.map((obj) => fields.map((f) => String(obj?.[f] ?? "")));

  sheetXml = replaceRowsInlineStrings(sheetXml, 2, matrix);

  const lastCol = colLetter(fields.length - 1);
  const dataRows = Math.max(1, matrix.length);
  const lastRow = 1 + dataRows;
  const ref = `A1:${lastCol}${lastRow}`;

  sheetXml = upsertDimension(sheetXml, ref);
  tableXml = updateTableRef(tableXml, ref);

  zip.file(fullSheetPath, sheetXml);
  zip.file(tablePath, tableXml);
}

// ---------- Normalization (Make-proofing) ----------

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
    const p = tryParseJson(v);
    if (p && typeof p === "object") v = p;
  }

  if (Array.isArray(v) && v.length === 1 && typeof v[0] === "object") v = v[0];

  if (v && typeof v === "object" && !Array.isArray(v)) return v;

  // Never fail hard: return blank object so we can still render lines/sublines/funds
  return {};
}

function normalizeRows(value) {
  const parsed = tryParseJson(value);
  if (parsed && Array.isArray(parsed)) return parsed;

  if (Array.isArray(value)) {
    if (value.length === 1 && typeof value[0] === "string") {
      const inner = tryParseJson(value[0]);
      if (inner && Array.isArray(inner)) return inner;
    }
    return value
      .map((item) => tryParseJson(item) ?? item)
      .filter((x) => x && typeof x === "object" && !Array.isArray(x));
  }

  return [];
}

// ---------- XML / ZIP helpers ----------

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

function replaceRowsInlineStrings(sheetXml, startRow, valuesMatrix) {
  const start = Number(startRow);

  sheetXml = sheetXml.replace(/<sheetData>([\s\S]*?)<\/sheetData>/i, (m, inner) => {
    const cleaned = inner.replace(/<row\b[^>]*\br="(\d+)"[\s\S]*?<\/row>/gi, (rowXml, rStr) => {
      const r = Number(rStr);
      return r >= start ? "" : rowXml;
    });

    const rowsXml = valuesMatrix.map((vals, i) => {
      const r = start + i;
      const cells = vals.map((v, idx) => makeInlineStrCell(colLetter(idx) + r, v)).join("");
      return `<row r="${r}">${cells}</row>`;
    });

    if (rowsXml.length === 0) {
      rowsXml.push(`<row r="${start}"></row>`);
    }

    return `<sheetData>${cleaned.trimEnd()}${rowsXml.join("")}</sheetData>`;
  });

  return sheetXml;
}

function makeInlineStrCell(r, text) {
  const t = escapeXml(text ?? "");
  return `<c r="${r}" t="inlineStr"><is><t>${t}</t></is></c>`;
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
