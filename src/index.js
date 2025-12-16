import JSZip from "jszip";

const TEMPLATE_KEY = "templates/final-report-template.xlsm";

// Must match your Excel table column order for Data_Scenario
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

export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return new Response("Method Not Allowed", { status: 405 });
    }

    // Auth
    const apiKeyHeader = request.headers.get("x-api-key");
    if (!apiKeyHeader) return new Response("Forbidden: missing x-api-key", { status: 403 });
    if (!env.API_KEY) return new Response("Forbidden: env.API_KEY not configured", { status: 500 });
    if (apiKeyHeader !== env.API_KEY) return new Response("Forbidden: api key mismatch", { status: 403 });

    // Parse JSON
    let payload;
    try {
      payload = await request.json();
    } catch {
      return new Response("Invalid JSON", { status: 400 });
    }

    const { scenarioId, scenarioName, scenario } = payload;
    if (!scenarioId) return new Response("Missing scenarioId", { status: 400 });
    if (!scenario || typeof scenario !== "object") return new Response("Missing scenario object", { status: 400 });

    // Load template
    const templateObj = await env.TEMPLATES_BUCKET.get(TEMPLATE_KEY);
    if (!templateObj) return new Response(`Template not found: ${TEMPLATE_KEY}`, { status: 500 });

    const templateBytes = new Uint8Array(await templateObj.arrayBuffer());

    // Unzip workbook
    const zip = await JSZip.loadAsync(templateBytes);

    // Find the worksheet file for sheet name "Data_Scenario"
    const workbookXml = await readText(zip, "xl/workbook.xml");
    const workbookRelsXml = await readText(zip, "xl/_rels/workbook.xml.rels");

    const sheetRid = findSheetRidByName(workbookXml, "Data_Scenario");
    const sheetPath = resolveWorkbookRidToTarget(workbookRelsXml, sheetRid); // "worksheets/sheetN.xml"
    const fullSheetPath = "xl/" + sheetPath;

    let sheetXml = await readText(zip, fullSheetPath);

    // Resolve table relationship -> table XML
    const sheetRelsPath = `xl/worksheets/_rels/${basename(sheetPath)}.rels`;
    const sheetRelsXml = await readText(zip, sheetRelsPath);
    const tableTarget = findFirstTableTarget(sheetRelsXml); // "../tables/tableX.xml"
    const tablePath = normalizePath("xl/worksheets", tableTarget); // "xl/tables/tableX.xml"

    let tableXml = await readText(zip, tablePath);

    // Build row values in the exact column order
    const rowValues = SCENARIO_FIELDS.map((f) => String(scenario[f] ?? ""));

    // Write row 2 starting at A2
    sheetXml = upsertRowInlineStrings(sheetXml, 2, rowValues);

    // Dimension A1:R2 (18 cols -> R)
    sheetXml = upsertDimension(sheetXml, "A1:R2");

    // Update the table range (header + one data row)
    tableXml = updateTableRef(tableXml, "A1:R2");

    // Save changes back into the zip
    zip.file(fullSheetPath, sheetXml);
    zip.file(tablePath, tableXml);

    // Re-zip
    const outBytes = await zip.generateAsync({ type: "uint8array" });

    // Save to reports bucket
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

    const fileUrl = `${env.R2_PUBLIC_BASE}/${outKey.split("/").map(encodeURIComponent).join("/")}`;

    return new Response(JSON.stringify({ fileUrl, fileName }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  },
};

// ---------- helpers ----------

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
  if (!m) throw new Error("No table relationship found for Data_Scenario sheet");
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

function upsertRowInlineStrings(sheetXml, rowNumber, values) {
  const rowRef = String(rowNumber);
  const cells = values.map((v, idx) => makeInlineStrCell(colLetter(idx) + rowRef, v)).join("");
  const newRow = `<row r="${rowRef}">${cells}</row>`;

  // Remove existing row 2 if present
  sheetXml = sheetXml.replace(new RegExp(`<row[^>]*r="${escapeReg(rowRef)}"[\\s\\S]*?<\\/row>`, "i"), "");

  // Append row inside sheetData
  sheetXml = sheetXml.replace(/<sheetData>([\s\S]*?)<\/sheetData>/i, (m, inner) => {
    const updated = inner.trimEnd() + newRow;
    return `<sheetData>${updated}</sheetData>`;
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
