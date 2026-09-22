import { NextRequest, NextResponse } from "next/server";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";
import JSZip from "jszip";

function normalizeString(s: unknown): string {
  return String(s ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function escapeXml(unsafe: unknown): string {
  return String(unsafe ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function colIndexToLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = "";
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function parseNumericValue(val: unknown): number {
  if (typeof val === "number") {
    return isNaN(val) ? 0 : val;
  }
  if (!val) return 0;

  let str = String(val).trim();
  let isNegative = false;
  if (str.startsWith("(") && str.endsWith(")")) {
    isNegative = true;
    str = str.slice(1, -1);
  }
  str = str.replace(/[$,\s]/g, "");
  const parsed = parseFloat(str);
  if (isNaN(parsed)) return 0;
  return isNegative ? -Math.abs(parsed) : parsed;
}

function formatDateValue(val: unknown): string {
  if (!val) return "";
  if (val instanceof Date) {
    return `${val.getMonth() + 1}/${val.getDate()}/${val.getFullYear()}`;
  }
  if (typeof val === "number" && val > 30000 && val < 60000) {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    return `${date.getUTCMonth() + 1}/${date.getUTCDate()}/${date.getUTCFullYear()}`;
  }
  return String(val).trim();
}

function dateToExcelSerial(dateVal: unknown): number {
  if (typeof dateVal === "number" && dateVal > 30000 && dateVal < 60000) {
    return Math.round(dateVal);
  }
  const str = String(dateVal || "").trim();
  const parts = str.split("/");
  if (parts.length === 3) {
    const m = parseInt(parts[0], 10) - 1;
    const d = parseInt(parts[1], 10);
    const y = parseInt(parts[2], 10);
    const utcMs = Date.UTC(y, m, d);
    return Math.round(utcMs / 86400000 + 25569);
  }
  const dt = new Date(str);
  if (!isNaN(dt.getTime())) {
    const utcMs = Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate());
    return Math.round(utcMs / 86400000 + 25569);
  }
  return 0;
}

function formatFilenameDate(serial: number): string {
  if (serial <= 0) return "";
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  const m = date.getUTCMonth() + 1;
  const d = String(date.getUTCDate()).padStart(2, "0");
  const y = String(date.getUTCFullYear()).slice(-2);
  return `${m}.${d}.${y}`;
}

function cleanLocationName(val: unknown): string {
  if (!val) return "";
  let str = String(val).trim();
  // Normalize Norfork typo if present
  str = str.replace(/\bnorfork\b/i, "Norfolk");
  // Do not remove "-b" / "-B" for Norfolk (preserve "Norfolk-B")
  if (/norfolk/i.test(str)) {
    if (/\s*-\s*b$/i.test(str)) {
      return str.replace(/\s*-\s*b$/i, "-B").trim();
    }
    if (/^norfolk$/i.test(str)) {
      return "Norfolk-B";
    }
    return str;
  }
  // Remove "-b", "-B", "-c", "-C", " - b", " - C" suffixes from location name
  return str.replace(/\s*-\s*[bc]$/i, "").trim();
}

function cleanStoreForMatching(val: unknown): string {
  if (!val) return "";
  let str = String(val).trim();
  str = str.replace(/\bnorfork\b/i, "Norfolk");
  // Remove trailing branch suffixes like "-b", "-c" (except Norfolk)
  if (!/norfolk/i.test(str)) {
    str = str.replace(/\s*-\s*[bc]$/i, "").trim();
  } else if (/\s*-\s*b$/i.test(str)) {
    str = str.replace(/\s*-\s*b$/i, "-B").trim();
  }
  // Remove leading store / unit / MOD number prefixes (e.g. "231 Kennesaw", "231 - Kennesaw")
  // but protect street numbers like "26 & Van Dyke" or "104th & Federal"
  str = str.replace(/^(?:mod\s*#?\s*)?\d+\s*[-_–—:]\s*/i, "").trim();
  str = str.replace(/^(?:mod\s*#?\s*)?\d+\s+(?![&]|st\b|nd\b|rd\b|th\b)/i, "").trim();
  return str;
}

function findBrinkForStore(
  storeName: string,
  brinkStoreData: Map<string, BrinkStoreData>,
  lookupTableauToBrink: Map<string, string>
): BrinkStoreData | undefined {
  const normRaw = normalizeString(cleanLocationName(storeName));
  const normClean = normalizeString(cleanStoreForMatching(storeName));

  // 1. Direct match with raw or cleaned name
  if (brinkStoreData.has(normRaw)) return brinkStoreData.get(normRaw);
  if (brinkStoreData.has(normClean)) return brinkStoreData.get(normClean);

  // 2. Lookup alias with raw or cleaned name
  const aliasRaw = lookupTableauToBrink.get(normRaw);
  if (aliasRaw) {
    const normAlias = normalizeString(cleanLocationName(aliasRaw));
    if (brinkStoreData.has(normAlias)) return brinkStoreData.get(normAlias);
  }

  const aliasClean = lookupTableauToBrink.get(normClean);
  if (aliasClean) {
    const normAlias = normalizeString(cleanLocationName(aliasClean));
    if (brinkStoreData.has(normAlias)) return brinkStoreData.get(normAlias);
  }

  // 3. Fallback: fuzzy/contains match across all Brink stores
  for (const [bKey, bVal] of brinkStoreData.entries()) {
    if (
      bKey === normRaw ||
      bKey === normClean ||
      (bKey.length >= 4 && (normClean.includes(bKey) || bKey.includes(normClean)))
    ) {
      return bVal;
    }
  }

  return undefined;
}

interface BrinkStoreData {
  locName: string;
  gross: number;
  mpf: number;
  surcharges: number;
  refunds: number;
  discounts: number;
  gcPromo: number;
  brinkGross: number; // Gross - MPF
  brinkNet: number;   // (Gross - MPF) - Refunds - Discounts + GCPromo - Surcharges
}

interface TableauRow {
  date: string;
  modNo: number | string;
  store: string;
  gross: number;
  net: number;
  franchisee: string;
}

// ----------------------------------------------------------------------
// EXTRACT BRINK STORE DATA FROM MASTER TEMPLATE
// ----------------------------------------------------------------------
function extractBrinkDataFromMaster(
  masterWb: XLSX.WorkBook,
  gcPromoMap: Map<string, number>
): Map<string, BrinkStoreData> {
  const map = new Map<string, BrinkStoreData>();
  const brinkSheet = masterWb.Sheets["Brink Sales Summary by Location"];
  if (!brinkSheet) return map;

  const rawRows: any[][] = XLSX.utils.sheet_to_json(brinkSheet, { header: 1 });

  // Data starts at row index 7 (row 8)
  for (let r = 7; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || !row[1]) continue;

    const rawLocName = String(row[1]).trim();
    if (normalizeString(rawLocName) === "total") continue;

    const locName = cleanLocationName(rawLocName);
    const gross = parseNumericValue(row[7]);
    const mpf = parseNumericValue(row[10]);
    const surcharges = parseNumericValue(row[11]);
    const refunds = parseNumericValue(row[12]);
    const discounts = parseNumericValue(row[13]);
    const gcPromo = gcPromoMap.get(normalizeString(locName)) || 0;

    const brinkGross = Math.round((gross - mpf) * 100) / 100;
    const brinkNet =
      Math.round(((gross - mpf) - refunds - discounts + gcPromo - surcharges) * 100) /
      100;

    map.set(normalizeString(locName), {
      locName,
      gross,
      mpf,
      surcharges,
      refunds,
      discounts,
      gcPromo,
      brinkGross,
      brinkNet,
    });
  }

  return map;
}

// ----------------------------------------------------------------------
// EXTRACT TABLEAU ROWS FROM MASTER TEMPLATE
// ----------------------------------------------------------------------
function extractTableauRowsFromMaster(masterWb: XLSX.WorkBook): TableauRow[] {
  const tabSheet = masterWb.Sheets["Tableau Data"];
  if (!tabSheet) return [];

  const rawRows: any[][] = XLSX.utils.sheet_to_json(tabSheet, { header: 1 });
  const rows: TableauRow[] = [];

  for (let r = 3; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || !row[2]) continue;

    rows.push({
      date: formatDateValue(row[0]),
      modNo: row[1] ?? 0,
      store: String(row[2]).trim(),
      gross: parseNumericValue(row[3]),
      net: parseNumericValue(row[4]),
      franchisee: String(row[5] ?? "").trim(),
    });
  }

  return rows;
}

// ----------------------------------------------------------------------
// RECONCILE VARIANCE (IF VARIANCE > 1, DEDUCT FROM TABLEAU NET_SALES)
// ----------------------------------------------------------------------
interface ReconcileResult {
  reconciledStoresCount: number;
  totalDeducted: number;
  reconciledStores: Map<string, { newTableauNet: number; variance: number }>;
}

function reconcileTableauVariance(
  validTableauRows: TableauRow[],
  brinkStoreData: Map<string, BrinkStoreData>,
  lookupTableauToBrink: Map<string, string>
): ReconcileResult {
  const tableauByStore = new Map<string, TableauRow[]>();
  for (const row of validTableauRows) {
    const storeKey = normalizeString(cleanStoreForMatching(row.store));
    if (!tableauByStore.has(storeKey)) {
      tableauByStore.set(storeKey, []);
    }
    tableauByStore.get(storeKey)!.push(row);
  }

  let reconciledStoresCount = 0;
  let totalDeducted = 0;
  const reconciledStores = new Map<string, { newTableauNet: number; variance: number }>();

  for (const [storeKey, storeRows] of tableauByStore.entries()) {
    const storeName = storeRows[0].store;
    const tableauNet =
      Math.round(storeRows.reduce((sum, r) => sum + r.net, 0) * 100) / 100;

    const bData = findBrinkForStore(storeName, brinkStoreData, lookupTableauToBrink);
    if (!bData) continue;

    // Variance: Tableau Net Sales - Brink Net Sales
    const variance = Math.round((tableauNet - bData.brinkNet) * 100) / 100;

    console.log(
      `[Reconcile] Store: ${storeName}, TableauNet: ${tableauNet}, BrinkNet: ${bData.brinkNet}, Variance: ${variance}`
    );

    // If variance > 0 (Tableau has more net sales than Brink), deduct the difference
    // from any Tableau Data row for that store so the variance becomes 0.
    // Threshold > 1 to avoid floating-point noise.
    if (variance > 1) {
      let remaining = variance;
      // Sort rows by net descending to safely deduct from highest net sales days
      const sortedRows = [...storeRows].sort((a, b) => b.net - a.net);

      for (const r of sortedRows) {
        if (remaining <= 0) break;
        const deduction = Math.min(r.net, remaining);
        r.net = Math.round((r.net - deduction) * 100) / 100;
        remaining = Math.round((remaining - deduction) * 100) / 100;
        totalDeducted += deduction;
      }

      reconciledStoresCount++;
      const newNet = Math.round((tableauNet - variance) * 100) / 100;
      reconciledStores.set(storeKey, { newTableauNet: newNet, variance });
      reconciledStores.set(normalizeString(cleanLocationName(storeName)), {
        newTableauNet: newNet,
        variance,
      });
    }
  }

  return {
    reconciledStoresCount,
    totalDeducted: Math.round(totalDeducted * 100) / 100,
    reconciledStores,
  };
}

// ----------------------------------------------------------------------
// SAFELY UPDATE SUMMARY SHEET — helpers
// ----------------------------------------------------------------------

/** Convert a column letter (A, B, ..., Z, AA, ...) to a 1-based index. */
function colLetterToIndex(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/**
 * Known default style indexes for columns in Summary (sheet2):
 * A: 2 (string)
 * B, C, M, N: 12 (number / currency)
 * D, E, F, G, H, I, J: 4 (blue fill, number format #,##0)
 * K, L: 11 (peach fill, number format #,##0)
 */
const DEFAULT_COLUMN_STYLES: Record<string, string> = {
  A: "2",
  B: "12",
  C: "12",
  D: "4",
  E: "4",
  F: "4",
  G: "4",
  H: "4",
  I: "4",
  J: "4",
  K: "11",
  L: "11",
  M: "12",
  N: "12",
};

/**
 * Update a cell in a row XML fragment while STRICTLY PRESERVING:
 * 1. Cell style attribute s="..." (fills, borders, fonts, number formats like #,##0)
 * 2. Formula tags (<f>...</f> or shared formula tags like <f t="shared" .../>)
 *
 * If newFormula is provided, sets/updates <f> (e.g. adding safe IFERROR fallback to Col I).
 * Otherwise preserves existing formula tag verbatim.
 */
function updateCellInRow(
  rowXml: string,
  cellRef: string,
  value: number,
  newFormula?: string
): string {
  const colLetter = cellRef.replace(/\d+/g, "");
  const colIdx = colLetterToIndex(colLetter);
  const defaultStyle = DEFAULT_COLUMN_STYLES[colLetter] || "4";

  // Match existing cell: either <c r="X" ...>...</c> or self-closing <c r="X" .../>
  const cellRegex = new RegExp(
    `<c\\s+r="${cellRef}"([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/c>)`,
    "i"
  );
  const match = rowXml.match(cellRegex);

  if (match) {
    const rawAttrs = match[1] || "";
    const innerContent = match[2] || "";

    // Keep all attributes (especially s="..."), remove t="..." since new value is numeric
    let cleanAttrs = rawAttrs.replace(/\s+t="[^"]*"/g, "");

    // Ensure s="..." is present
    if (!/\bs="\d+"/.test(cleanAttrs)) {
      cleanAttrs = ` s="${defaultStyle}"` + cleanAttrs;
    }

    // Determine formula
    let formulaTag = "";
    if (newFormula !== undefined) {
      formulaTag = `<f>${newFormula}</f>`;
    } else {
      // Preserve existing <f...>...</f> or <f.../> tag
      const fMatch = innerContent.match(/<f(?:\s+[^>]*)?(?:>[\s\S]*?<\/f>|\/>)/i);
      if (fMatch) {
        formulaTag = fMatch[0];
      }
    }

    const replacementCell = `<c r="${cellRef}"${cleanAttrs}>${formulaTag}<v>${value}</v></c>`;
    return rowXml.replace(match[0], replacementCell);
  }

  // Cell doesn't exist in the row – insert in sorted column order
  const styleAttr = ` s="${defaultStyle}"`;
  const formulaTag = newFormula ? `<f>${newFormula}</f>` : "";
  const newCellXml = `<c r="${cellRef}"${styleAttr}>${formulaTag}<v>${value}</v></c>`;

  const cellTagRe = /<c\s+r="([A-Z]+)\d+"(?:\s+[^>]*)?(?:\/>|>[\s\S]*?<\/c>)/gi;
  let m: RegExpExecArray | null;
  let insertIdx = -1;

  while ((m = cellTagRe.exec(rowXml)) !== null) {
    if (colLetterToIndex(m[1]) > colIdx) {
      insertIdx = m.index;
      break;
    }
  }

  if (insertIdx !== -1) {
    return rowXml.slice(0, insertIdx) + newCellXml + rowXml.slice(insertIdx);
  }
  return rowXml.replace("</row>", newCellXml + "</row>");
}

async function updateReconciledCellsInSummary(
  zip: JSZip,
  masterWb: XLSX.WorkBook,
  reconciledStores: Map<string, { newTableauNet: number; variance: number }>,
  gcPromoMap: Map<string, number>,
  brinkStoreData: Map<string, BrinkStoreData>,
  lookupTableauToBrink: Map<string, string>
): Promise<void> {
  const s2File = zip.file("xl/worksheets/sheet2.xml");
  if (!s2File) return;

  let s2Xml = await s2File.async("text");
  const summarySheet = masterWb.Sheets["Summary"];
  if (!summarySheet) return;

  for (let r = 4; r <= 85; r++) {
    const cellA = summarySheet[`A${r}`];
    if (!cellA || !cellA.v) continue;

    const storeName = String(cellA.v).trim();
    const normClean = normalizeString(cleanStoreForMatching(storeName));
    const normRaw = normalizeString(cleanLocationName(storeName));

    // ── 1. Reconciled tableau net (col C) & variance (col L) ──────────────
    let reconMatch = reconciledStores.get(normClean) || reconciledStores.get(normRaw);
    if (!reconMatch) {
      for (const [rKey, rVal] of reconciledStores.entries()) {
        if (
          rKey === normClean ||
          rKey === normRaw ||
          (rKey.length >= 5 && (normClean.includes(rKey) || rKey.includes(normClean)))
        ) {
          reconMatch = rVal;
          break;
        }
      }
    }

    // ── 2. Look up GC Promo value for this store ───────────────────────────
    let gcPromo = gcPromoMap.get(normClean) ?? gcPromoMap.get(normRaw) ?? 0;
    if (gcPromo === 0) {
      // Try alias via Lookup sheet
      const alias = lookupTableauToBrink.get(normClean) || lookupTableauToBrink.get(normRaw);
      if (alias) {
        gcPromo = gcPromoMap.get(normalizeString(alias)) ?? 0;
      }
    }
    if (gcPromo === 0) {
      // Fuzzy search in gcPromoMap
      for (const [k, v] of gcPromoMap.entries()) {
        if (k.length >= 4 && (normClean.includes(k) || k.includes(normClean))) {
          gcPromo = v;
          break;
        }
      }
    }

    // ── 3. Look up Brink data for this store (for Net Sales computation) ───
    const bData = findBrinkForStore(storeName, brinkStoreData, lookupTableauToBrink);

    const rowRegex = new RegExp(`(<row\\s+r="${r}"(?:\\s[^>]*)?>)([\\s\\S]*?)(</row>)`, "i");
    const rowMatch = s2Xml.match(rowRegex);
    if (!rowMatch) continue;

    let rowXml = rowMatch[0];

    // ── 3a. Update Brink columns cached values while PRESERVING formulas & styles ──
    // Formula E4: VLOOKUP(B:H, 7=GrossSales) - D4(Surcharges)
    // Formula F4: VLOOKUP(B:N, 12=Refunds)
    // Formula G4: VLOOKUP(B:N, 11=Surcharges)
    // Formula H4: VLOOKUP(B:N, 13=Discounts)
    if (bData) {
      rowXml = updateCellInRow(rowXml, `E${r}`, Math.round(bData.brinkGross * 100) / 100);
      rowXml = updateCellInRow(rowXml, `F${r}`, Math.round(bData.refunds * 100) / 100);
      rowXml = updateCellInRow(rowXml, `G${r}`, Math.round(bData.surcharges * 100) / 100);
      rowXml = updateCellInRow(rowXml, `H${r}`, Math.round(bData.discounts * 100) / 100);
    }

    // ── 3b. Update GC Promos (col I) with safe formula & calculated value ──
    // Original template formula returns #N/A when store/alias is missing from GC Promo sheet.
    // We add safe IFERROR(..., 0) fallback so it evaluates to 0 instead of #N/A.
    const safeGcPromoFormula = `IFERROR(VLOOKUP(A${r},'GC Promo'!A:F,6,0),IFERROR(VLOOKUP(VLOOKUP(A${r},Lookup!A:B,2,0),'GC Promo'!A:F,6,0),0))`;
    rowXml = updateCellInRow(rowXml, `I${r}`, gcPromo, safeGcPromoFormula);

    // ── 3c. Update Net Sales (col J) cached value while PRESERVING formula & style ──
    // Formula J4: =E4-F4-H4+I4-G4
    if (bData) {
      const netSales = Math.round(
        (bData.brinkGross - bData.refunds - bData.discounts + gcPromo - bData.surcharges) * 100
      ) / 100;
      rowXml = updateCellInRow(rowXml, `J${r}`, netSales);
    }

    // ── 3d. Reconciled tableau net (col C) & variance (col L) ─────────────
    if (reconMatch) {
      rowXml = updateCellInRow(rowXml, `C${r}`, reconMatch.newTableauNet);
      rowXml = updateCellInRow(rowXml, `L${r}`, 0);
    }

    s2Xml = s2Xml.replace(rowMatch[0], rowXml);
  }

  zip.file("xl/worksheets/sheet2.xml", s2Xml);
}

// ----------------------------------------------------------------------
// WRITE TABLEAU ROWS TO SHEET6.XML & TABLE1.XML
// ----------------------------------------------------------------------
async function writeTableauRowsToZip(
  zip: JSZip,
  validTableauRows: TableauRow[]
): Promise<void> {
  const s6File = zip.file("xl/worksheets/sheet6.xml");
  if (!s6File) {
    throw new Error("Could not find sheet6.xml (Tableau Data) in template.");
  }

  const originalS6Xml = await s6File.async("text");
  const row4StartIdx = originalS6Xml.indexOf('<row r="4"');
  const sheetDataEndIdx = originalS6Xml.indexOf("</sheetData>");

  if (row4StartIdx === -1 || sheetDataEndIdx === -1) {
    throw new Error("Could not locate data section in sheet6.xml.");
  }

  const headerXml = originalS6Xml.slice(0, row4StartIdx);
  const footerXml = originalS6Xml.slice(sheetDataEndIdx);

  let newRowsXml = "";
  validTableauRows.forEach((row, idx) => {
    const rNum = 4 + idx;
    newRowsXml += `<row r="${rNum}" spans="1:8" x14ac:dyDescent="0.25">`;
    newRowsXml += `<c r="A${rNum}" s="18" t="str"><v>${escapeXml(row.date)}</v></c>`;
    newRowsXml += `<c r="B${rNum}"><v>${row.modNo}</v></c>`;
    newRowsXml += `<c r="C${rNum}" t="str"><v>${escapeXml(row.store)}</v></c>`;
    newRowsXml += `<c r="D${rNum}" s="46"><v>${row.gross}</v></c>`;
    newRowsXml += `<c r="E${rNum}" s="46"><v>${row.net}</v></c>`;
    newRowsXml += `<c r="F${rNum}" s="16" t="str"><f>VLOOKUP(Table1[[#This Row],[Store]],'Rates from Tracker'!A:I,3,0)</f><v>${escapeXml(row.franchisee)}</v></c>`;
    newRowsXml += `</row>`;
  });

  const lastRow = 3 + validTableauRows.length;
  let updatedS6Xml = headerXml + newRowsXml + footerXml;
  updatedS6Xml = updatedS6Xml.replace(
    /<dimension ref="[^"]*"/,
    `<dimension ref="A1:H${lastRow}"`
  );

  zip.file("xl/worksheets/sheet6.xml", updatedS6Xml);

  const t1File = zip.file("xl/tables/table1.xml");
  if (t1File) {
    let t1Xml = await t1File.async("text");
    t1Xml = t1Xml.replace(/ref="A3:F\d+"/, `ref="A3:F${lastRow}"`);
    zip.file("xl/tables/table1.xml", t1Xml);
  }
}



// ----------------------------------------------------------------------
// PARSE UPLOADED TABLEAU FILE
// ----------------------------------------------------------------------
function parseUploadedTableauBuffer(
  tableauBuffer: Buffer,
  locationToFranchisee: Map<string, string>
): TableauRow[] {
  const tableauWb = XLSX.read(tableauBuffer, { type: "buffer" });
  const chosenSheetName =
    tableauWb.SheetNames.find((name) => /tableau/i.test(name)) ||
    tableauWb.SheetNames.find((name) => /data/i.test(name)) ||
    tableauWb.SheetNames.find((name) => /export/i.test(name)) ||
    tableauWb.SheetNames.find((name) => /sales/i.test(name)) ||
    tableauWb.SheetNames[0];

  const rawRows: any[][] = XLSX.utils.sheet_to_json(
    tableauWb.Sheets[chosenSheetName],
    { header: 1 }
  );

  if (rawRows.length === 0) {
    throw new Error("The uploaded Tableau sheet is empty.");
  }

  let headerRowIdx = 0;
  for (let r = 0; r < Math.min(10, rawRows.length); r++) {
    const rStr = (rawRows[r] || []).map(normalizeString).join(" ");
    if (
      rStr.includes("store") &&
      (rStr.includes("gross") || rStr.includes("date") || rStr.includes("mod"))
    ) {
      headerRowIdx = r;
      break;
    }
  }

  const headerRow = rawRows[headerRowIdx] || [];
  let dateCol = -1;
  let modCol = -1;
  let storeCol = -1;
  let grossCol = -1;
  let netCol = -1;

  for (let c = 0; c < headerRow.length; c++) {
    const norm = normalizeString(headerRow[c]);
    if (norm.includes("date") && dateCol === -1) dateCol = c;
    else if ((norm.includes("mod") || norm.includes("unit")) && modCol === -1)
      modCol = c;
    else if (norm.includes("store") && storeCol === -1) storeCol = c;
    else if (norm.includes("gross") && grossCol === -1) grossCol = c;
    else if (norm.includes("net") && netCol === -1) netCol = c;
  }

  if (dateCol === -1 && headerRow.length >= 5) dateCol = 0;
  if (modCol === -1 && headerRow.length >= 5) modCol = 1;
  if (storeCol === -1 && headerRow.length >= 5) storeCol = 2;
  if (grossCol === -1 && headerRow.length >= 5) grossCol = 3;
  if (netCol === -1 && headerRow.length >= 5) netCol = 4;

  if (storeCol === -1) {
    throw new Error(
      "Could not identify the 'Store' column in the uploaded Tableau file."
    );
  }

  const validTableauRows: TableauRow[] = [];

  for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0) continue;

    const storeName = String(row[storeCol] ?? "").trim();
    if (!storeName || normalizeString(storeName) === "store") continue;

    const dateVal = dateCol !== -1 ? formatDateValue(row[dateCol]) : "";
    const modVal = modCol !== -1 ? parseNumericValue(row[modCol]) : 0;
    const grossVal = grossCol !== -1 ? parseNumericValue(row[grossCol]) : 0;
    const netVal = netCol !== -1 ? parseNumericValue(row[netCol]) : 0;
    const fran = locationToFranchisee.get(normalizeString(storeName)) || "";

    validTableauRows.push({
      date: dateVal,
      modNo: modVal,
      store: storeName,
      gross: Math.round(grossVal * 100) / 100,
      net: Math.round(netVal * 100) / 100,
      franchisee: fran,
    });
  }

  if (validTableauRows.length === 0) {
    throw new Error("No data rows found in the uploaded Tableau file.");
  }

  return validTableauRows;
}

// ----------------------------------------------------------------------
// UPDATE BRINK SALES SUMMARY (Sheet #3)
// ----------------------------------------------------------------------
async function processBrinkUpdate(
  zip: JSZip,
  masterBuffer: Buffer,
  brinkBuffer: Buffer,
  gcPromoMap: Map<string, number>
): Promise<{
  rowsCount: number;
  matchedColsCount: number;
  brinkStoreData: Map<string, BrinkStoreData>;
}> {
  const masterWb = XLSX.read(masterBuffer, { type: "buffer" });
  const targetSheet = masterWb.Sheets["Brink Sales Summary by Location"];

  if (!targetSheet) {
    throw new Error("Sheet 'Brink Sales Summary by Location' not found.");
  }

  const targetJson: any[][] = XLSX.utils.sheet_to_json(targetSheet, {
    header: 1,
  });
  const templateRow5 = targetJson[5] || [];
  const templateRow6 = targetJson[6] || [];

  const totalTargetColumns = 338;
  const targetCols: { colIndex: number; name: string }[] = [];
  const targetNormMap = new Map<string, number>();

  let currentCat = "";
  for (let c = 0; c < totalTargetColumns; c++) {
    if (templateRow5[c]) currentCat = String(templateRow5[c]).trim();
    const subName =
      templateRow6[c] !== undefined && templateRow6[c] !== null
        ? String(templateRow6[c]).trim()
        : "";

    let colName = subName;
    if (c === 0) colName = "Location Code";
    else if (c === 1) colName = "Location Name";
    else if (!colName && templateRow5[c])
      colName = String(templateRow5[c]).trim();
    if (!colName) colName = `Column_${c}`;

    targetCols.push({ colIndex: c, name: colName });

    const norm = normalizeString(colName);
    if (norm && !targetNormMap.has(norm)) {
      targetNormMap.set(norm, c);
    }

    if (subName && currentCat && subName !== currentCat) {
      const composite = normalizeString(`${currentCat} ${subName}`);
      if (!targetNormMap.has(composite)) {
        targetNormMap.set(composite, c);
      }
    }
  }

  const aliases: Record<string, string[]> = {
    locationcode: [
      "location_code",
      "locationid",
      "location_id",
      "storecode",
      "store_code",
      "store#",
      "storeid",
      "storenumber",
      "store_no",
      "unit",
      "unit#",
    ],
    locationname: [
      "location_name",
      "storename",
      "store_name",
      "location",
      "store",
      "store_description",
    ],
    taxablesales: ["taxable_sales", "taxablesale", "taxablesalesamount"],
    nontaxablesales: [
      "non_taxable_sales",
      "nontaxablesale",
      "nontaxsales",
      "non_tax_sales",
    ],
    netsales: ["net_sales", "netsale", "netsalesamount"],
    grosssales: ["gross_sales", "grosssale", "grosssalesamount"],
    tax: ["taxes", "salestax", "sales_tax", "taxamount"],
    mpftaxremitted: ["mpf_tax_remitted", "mpftax", "mpf_tax", "mpf_remitted"],
    surcharges: ["surcharge", "surcharges"],
    refunds: ["refund", "refunds", "refundamount"],
    discounts: ["discount", "discounts", "totaldiscounts"],
    voided: ["voids", "void", "voided_sales", "voidamount"],
    tips: ["tip", "tips", "total_tips", "creditcardtips"],
    paidin: ["paid_in"],
    paidout: ["paid_out"],
    guestcount: [
      "guest_count",
      "guests",
      "guestscount",
      "covercount",
      "covers",
    ],
    ordercount: ["order_count", "orders", "orderscount"],
  };

  for (const [targetKey, aliasList] of Object.entries(aliases)) {
    const targetIdx = targetNormMap.get(normalizeString(targetKey));
    if (targetIdx !== undefined) {
      for (const alias of aliasList) {
        const aNorm = normalizeString(alias);
        if (!targetNormMap.has(aNorm)) {
          targetNormMap.set(aNorm, targetIdx);
        }
      }
    }
  }

  const brinkWb = XLSX.read(brinkBuffer, { type: "buffer" });
  const chosenSheetName =
    brinkWb.SheetNames.find((name) =>
      /sales.*summary|summary.*location|brink|sales/i.test(name)
    ) || brinkWb.SheetNames[0];

  const rawRows: any[][] = XLSX.utils.sheet_to_json(
    brinkWb.Sheets[chosenSheetName],
    { header: 1 }
  );

  if (rawRows.length === 0) {
    throw new Error("The uploaded Brink sales summary sheet is empty.");
  }

  let isBrinkTwoLevel = false;
  let categoryRowIdx = -1;
  let subHeaderRowIdx = -1;
  let dataStartRowIdx = -1;

  for (let r = 0; r < Math.min(10, rawRows.length - 1); r++) {
    const rRow = rawRows[r] || [];
    const nextRow = rawRows[r + 1] || [];

    const rHasLoc = rRow.some((cell) =>
      /location code|location name/i.test(String(cell || ""))
    );
    const nextHasSales = nextRow.some((cell) =>
      /taxable sales|net sales|gross sales/i.test(String(cell || ""))
    );

    if (rHasLoc && nextHasSales) {
      isBrinkTwoLevel = true;
      categoryRowIdx = r;
      subHeaderRowIdx = r + 1;
      dataStartRowIdx = r + 2;
      break;
    }
  }

  const uploadedHeaders: string[] = [];

  if (isBrinkTwoLevel) {
    const catRow = rawRows[categoryRowIdx] || [];
    const subRow = rawRows[subHeaderRowIdx] || [];
    const maxCols = Math.max(catRow.length, subRow.length);

    for (let c = 0; c < maxCols; c++) {
      const sub = String(subRow[c] ?? "").trim();
      const cat = String(catRow[c] ?? "").trim();

      if (c === 0 && !sub && /location code/i.test(cat)) {
        uploadedHeaders[c] = "Location Code";
      } else if (c === 1 && !sub && /location name/i.test(cat)) {
        uploadedHeaders[c] = "Location Name";
      } else {
        uploadedHeaders[c] = sub || cat || "";
      }
    }
  } else {
    let bestRowIdx = 0;
    let maxMatches = 0;

    for (let r = 0; r < Math.min(15, rawRows.length); r++) {
      const row = rawRows[r] || [];
      let matches = 0;
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (val !== undefined && val !== null && String(val).trim() !== "") {
          const norm = normalizeString(val);
          if (targetNormMap.has(norm)) {
            matches++;
          }
        }
      }
      if (matches > maxMatches) {
        maxMatches = matches;
        bestRowIdx = r;
      }
    }

    const hRow = rawRows[bestRowIdx] || [];
    for (let c = 0; c < hRow.length; c++) {
      uploadedHeaders[c] = String(hRow[c] ?? "").trim();
    }
    dataStartRowIdx = bestRowIdx + 1;
  }

  const colMapping: { targetCol: number; uploadedCol: number }[] = [];

  uploadedHeaders.forEach((header, uploadedColIdx) => {
    if (!header) return;
    const norm = normalizeString(header);
    if (targetNormMap.has(norm)) {
      const targetCol = targetNormMap.get(norm)!;
      if (!colMapping.some((m) => m.targetCol === targetCol)) {
        colMapping.push({ targetCol, uploadedCol: uploadedColIdx });
      }
    }
  });

  const validDataRows: any[][] = [];
  for (let r = dataStartRowIdx; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.length === 0) continue;

    const hasContent = row.some(
      (cell) =>
        cell !== undefined && cell !== null && String(cell).trim() !== ""
    );
    if (!hasContent) continue;

    const isTotal = row.some((cell) => {
      const s = String(cell || "").trim().toLowerCase();
      return s === "total" || s === "totals" || s === "grand total";
    });
    if (isTotal) continue;

    const matchCount = row.filter((cell) =>
      targetNormMap.has(normalizeString(cell))
    ).length;
    if (matchCount >= 3) continue;

    validDataRows.push(row);
  }

  if (validDataRows.length === 0) {
    throw new Error("No location sales data rows found in the Brink file.");
  }

  const sheet3File = zip.file("xl/worksheets/sheet3.xml");
  if (!sheet3File) {
    throw new Error("Could not find sheet3.xml in master archive.");
  }

  const originalSheet3Xml = await sheet3File.async("text");
  const row8StartIdx = originalSheet3Xml.indexOf('<row r="8"');
  const sheetDataEndIdx = originalSheet3Xml.indexOf("</sheetData>");

  if (row8StartIdx === -1 || sheetDataEndIdx === -1) {
    throw new Error("Could not locate data section in sheet3.xml.");
  }

  const headerXml = originalSheet3Xml.slice(0, row8StartIdx);
  const footerXml = originalSheet3Xml.slice(sheetDataEndIdx);

  const columnSums = new Array<number>(totalTargetColumns).fill(0);
  let newRowsXml = "";

  const brinkStoreData = new Map<string, BrinkStoreData>();

  validDataRows.forEach((uploadedRow, rowIdx) => {
    const targetRowNumber = 8 + rowIdx;
    newRowsXml += `<row r="${targetRowNumber}" ht="18" customHeight="1">`;

    let locName = "";
    let gross = 0;
    let mpf = 0;
    let surcharges = 0;
    let refunds = 0;
    let discounts = 0;

    for (let c = 0; c < totalTargetColumns; c++) {
      const colLetter = colIndexToLetter(c);
      const cellRef = `${colLetter}${targetRowNumber}`;
      const mapping = colMapping.find((m) => m.targetCol === c);

      if (c === 0) {
        let locCode = "";
        if (mapping) {
          locCode = String(uploadedRow[mapping.uploadedCol] ?? "").trim();
        } else if (uploadedRow[0] !== undefined) {
          locCode = String(uploadedRow[0]).trim();
        }
        // Use t="str" (formula-string) so VLOOKUP in Summary sheet can match against this cell
        newRowsXml += `<c r="${cellRef}" t="str"><v>${escapeXml(locCode)}</v></c>`;
      } else if (c === 1) {
        if (mapping) {
          locName = cleanLocationName(uploadedRow[mapping.uploadedCol]);
        } else if (uploadedRow[1] !== undefined) {
          locName = cleanLocationName(uploadedRow[1]);
        }
        // Use t="str" so VLOOKUP(A4,'Brink Sales Summary by Location'!B:H,...) can find the store name
        newRowsXml += `<c r="${cellRef}" t="str"><v>${escapeXml(locName)}</v></c>`;
      } else if (c === 4 || c === 5 || c === 9) {
        newRowsXml += `<c r="${cellRef}"><v>0</v></c>`;
      } else {
        let numVal = 0;
        if (mapping) {
          numVal = parseNumericValue(uploadedRow[mapping.uploadedCol]);
        }
        columnSums[c] += numVal;
        const rounded = Math.round(numVal * 100) / 100;
        newRowsXml += `<c r="${cellRef}"><v>${rounded}</v></c>`;

        if (c === 7) gross = rounded;
        else if (c === 10) mpf = rounded;
        else if (c === 11) surcharges = rounded;
        else if (c === 12) refunds = rounded;
        else if (c === 13) discounts = rounded;
      }
    }
    newRowsXml += "</row>";

    if (locName) {
      const gcPromo = gcPromoMap.get(normalizeString(locName)) || 0;
      const brinkGross = Math.round((gross - mpf) * 100) / 100;
      const brinkNet =
        Math.round(((gross - mpf) - refunds - discounts + gcPromo - surcharges) * 100) /
        100;

      brinkStoreData.set(normalizeString(locName), {
        locName,
        gross,
        mpf,
        surcharges,
        refunds,
        discounts,
        gcPromo,
        brinkGross,
        brinkNet,
      });
    }
  });

  const totalRowNumber = 8 + validDataRows.length;
  newRowsXml += `<row r="${totalRowNumber}" ht="18" customHeight="1">`;
  newRowsXml += `<c r="B${totalRowNumber}" t="str"><v>Total</v></c>`;

  for (let c = 2; c < totalTargetColumns; c++) {
    const colLetter = colIndexToLetter(c);
    const cellRef = `${colLetter}${totalRowNumber}`;

    if (c === 4 || c === 5 || c === 9) {
      newRowsXml += `<c r="${cellRef}"><v>0</v></c>`;
    } else {
      const roundedSum = Math.round(columnSums[c] * 100) / 100;
      newRowsXml += `<c r="${cellRef}"><v>${roundedSum}</v></c>`;
    }
  }
  newRowsXml += "</row>";

  let updatedSheetXml = headerXml + newRowsXml + footerXml;
  updatedSheetXml = updatedSheetXml.replace(
    /<dimension ref="[^"]*"/,
    `<dimension ref="A1:LZ${totalRowNumber}"`
  );
  updatedSheetXml = updatedSheetXml.replace(
    /<autoFilter ref="[^"]*"/,
    `<autoFilter ref="A7:LZ${totalRowNumber}"`
  );

  zip.file("xl/worksheets/sheet3.xml", updatedSheetXml);

  return {
    rowsCount: validDataRows.length,
    matchedColsCount: colMapping.length,
    brinkStoreData,
  };
}

// ----------------------------------------------------------------------
// REMOVE STALE CALCULATION CHAIN TO PREVENT EXCEL REPAIR WARNINGS
// ----------------------------------------------------------------------
async function cleanCalculationChain(zip: JSZip): Promise<void> {
  // 1. Remove xl/calcChain.xml
  zip.remove("xl/calcChain.xml");

  // 2. Remove relationship in xl/_rels/workbook.xml.rels
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (relsFile) {
    let relsXml = await relsFile.async("text");
    relsXml = relsXml.replace(
      /<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/g,
      ""
    );
    zip.file("xl/_rels/workbook.xml.rels", relsXml);
  }

  // 3. Remove override in [Content_Types].xml
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ctXml = await ctFile.async("text");
    ctXml = ctXml.replace(
      /<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g,
      ""
    );
    zip.file("[Content_Types].xml", ctXml);
  }

  // 4. Force full calculation on load in xl/workbook.xml
  const wbFile = zip.file("xl/workbook.xml");
  if (wbFile) {
    let wbXml = await wbFile.async("text");
    if (wbXml.includes("<calcPr")) {
      wbXml = wbXml.replace(/<calcPr\b([^>]*?)(?:\/>|>[\s\S]*?<\/calcPr>)/, (match, attrs) => {
        let cleanAttrs = attrs.replace(/\s*\/$/, "").trim();
        if (!cleanAttrs.includes("fullCalcOnLoad")) {
          cleanAttrs += ' fullCalcOnLoad="1"';
        } else {
          cleanAttrs = cleanAttrs.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"');
        }
        return `<calcPr ${cleanAttrs}/>`;
      });
    }
    zip.file("xl/workbook.xml", wbXml);
  }
}

// ----------------------------------------------------------------------
// FORCE PIVOT TABLES TO REFRESH FROM SOURCE DATA ON FILE OPEN
// Without this, the pivot cache has stale Net_Sales values and the
// Summary sheet Column C (from PivotTable2) shows old numbers,
// making Variance (Net Sales) non-zero even though we already
// adjusted the Tableau Data sheet.
// ----------------------------------------------------------------------
async function enablePivotRefreshOnLoad(zip: JSZip): Promise<void> {
  const allFiles = Object.keys(zip.files);
  const pivotCacheFiles = allFiles.filter((f) =>
    /xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(f)
  );

  for (const pcFile of pivotCacheFiles) {
    const file = zip.file(pcFile);
    if (!file) continue;
    let xml = await file.async("text");

    if (!xml.includes("refreshOnLoad")) {
      // Insert refreshOnLoad="1" into the opening <pivotCacheDefinition> tag
      xml = xml.replace(
        "<pivotCacheDefinition ",
        '<pivotCacheDefinition refreshOnLoad="1" '
      );
      console.log(`[Pivot] Added refreshOnLoad to ${pcFile}`);
    } else {
      // Ensure it is set to "1"
      xml = xml.replace(/refreshOnLoad="[^"]*"/, 'refreshOnLoad="1"');
    }

    zip.file(pcFile, xml);
  }
}

// ----------------------------------------------------------------------
// MAIN API ROUTE HANDLER
// ----------------------------------------------------------------------
export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();

    const brinkFile = formData.get("brinkFile") as File | null;
    const tableauFile = formData.get("tableauFile") as File | null;
    const genericFile = formData.get("file") as File | null;
    const requestedTarget = (formData.get("targetType") as string) || "auto";

    if (!brinkFile && !tableauFile && !genericFile) {
      return NextResponse.json(
        { error: "Please upload at least one Excel file to update." },
        { status: 400 }
      );
    }

    const masterPath = path.join(
      process.cwd(),
      "Franchise Sales 8.10.26 - 8.16.26.xlsx"
    );
    const masterBuffer = await fs.promises.readFile(masterPath);
    const masterWb = XLSX.read(masterBuffer, { type: "buffer" });
    const zip = await JSZip.loadAsync(masterBuffer);

    // 1. Build Location -> Franchisee mapping from Rates from Tracker
    const ratesSheet = masterWb.Sheets["Rates from Tracker"];
    const ratesData: any[][] = XLSX.utils.sheet_to_json(ratesSheet, {
      header: 1,
    });
    const locationToFranchisee = new Map<string, string>();
    for (let r = 4; r < ratesData.length; r++) {
      const loc = ratesData[r][0];
      const debitLoc = ratesData[r][1];
      const fran = ratesData[r][2];
      if (loc && fran) {
        locationToFranchisee.set(normalizeString(loc), String(fran).trim());
      }
      if (debitLoc && fran) {
        locationToFranchisee.set(normalizeString(debitLoc), String(fran).trim());
      }
    }

    // 2. Build Lookup (Tableau Name -> Brink Name)
    const lookupTableauToBrink = new Map<string, string>();
    const lookupSheet = masterWb.Sheets["Lookup"];
    if (lookupSheet) {
      const lookupData: any[][] = XLSX.utils.sheet_to_json(lookupSheet, {
        header: 1,
      });
      for (let r = 1; r < lookupData.length; r++) {
        const tabName = lookupData[r][0];
        const brinkName = lookupData[r][1];
        if (tabName && brinkName) {
          lookupTableauToBrink.set(
            normalizeString(tabName),
            normalizeString(brinkName)
          );
          const fran = locationToFranchisee.get(normalizeString(brinkName));
          if (fran && !locationToFranchisee.has(normalizeString(tabName))) {
            locationToFranchisee.set(normalizeString(tabName), fran);
          }
        }
      }
    }

    // 3. Build GC Promo Map
    const gcPromoMap = new Map<string, number>();
    const gcSheet = masterWb.Sheets["GC Promo"];
    if (gcSheet) {
      const gcData: any[][] = XLSX.utils.sheet_to_json(gcSheet, { header: 1 });
      for (let r = 3; r < gcData.length; r++) {
        const row = gcData[r];
        if (row && row[0]) {
          gcPromoMap.set(normalizeString(row[0]), parseNumericValue(row[5]));
        }
      }
    }

    let updatedBrinkRows = 0;
    let updatedTableauRows = 0;
    const updatedSheetsList: string[] = [];

    let brinkStoreData: Map<string, BrinkStoreData> | null = null;
    let tableauRows: TableauRow[] | null = null;

    // Detect if single generic file is Tableau or Brink
    const detectIsTableau = (rawRows: any[][]) => {
      for (let r = 0; r < Math.min(5, rawRows.length); r++) {
        const rowStr = (rawRows[r] || [])
          .map((c) => normalizeString(c))
          .join(" ");
        if (
          (rowStr.includes("store") || rowStr.includes("date")) &&
          (rowStr.includes("gross") ||
            rowStr.includes("mod") ||
            rowStr.includes("net")) &&
          !rowStr.includes("tenders") &&
          !rowStr.includes("discounts")
        ) {
          return true;
        }
      }
      return false;
    };

    // Enforce both Brink and Tableau files in Step 1
    if (!brinkFile || !tableauFile) {
      return NextResponse.json(
        {
          error:
            "Both Brink Sales Summary and Tableau Data files must be uploaded together.",
        },
        { status: 400 }
      );
    }

    // A. Process Brink file
    const brinkBuf = Buffer.from(await brinkFile.arrayBuffer());
    const res = await processBrinkUpdate(zip, masterBuffer, brinkBuf, gcPromoMap);
    updatedBrinkRows = res.rowsCount;
    brinkStoreData = res.brinkStoreData;
    updatedSheetsList.push(`Brink Sales Summary (#3: ${res.rowsCount} stores)`);

    // B. Process Tableau file
    const tabBuf = Buffer.from(await tableauFile.arrayBuffer());
    tableauRows = parseUploadedTableauBuffer(tabBuf, locationToFranchisee);
    updatedTableauRows = tableauRows.length;
    updatedSheetsList.push(`Tableau Data (#6: ${tableauRows.length} records)`);

    // Ensure we have both datasets available for reconciliation and summary update
    if (!brinkStoreData) {
      brinkStoreData = extractBrinkDataFromMaster(masterWb, gcPromoMap);
    }
    if (!tableauRows) {
      tableauRows = extractTableauRowsFromMaster(masterWb);
    }

    // D. RECONCILE VARIANCE:
    // If any store has variance > 1 in Variance Net Sales, deduct that variance from Tableau Data sheet
    const varianceResult = reconcileTableauVariance(
      tableauRows,
      brinkStoreData,
      lookupTableauToBrink
    );

    // If Tableau was uploaded OR variance adjustments were made, write sheet6.xml
    if (tableauFile || (genericFile && updatedTableauRows > 0) || varianceResult.reconciledStoresCount > 0) {
      await writeTableauRowsToZip(zip, tableauRows);
      if (varianceResult.reconciledStoresCount > 0 && !updatedSheetsList.some(s => s.includes("Tableau Data"))) {
        updatedSheetsList.push(`Tableau Data (Reconciled ${varianceResult.reconciledStoresCount} stores)`);
      }
    }

    // E. Update Summary sheet:
    //    - Col C (Tableau Net Sales) for reconciled stores
    //    - Col I (GC Promos) with direct values — bypasses broken VLOOKUP
    //    - Col J (Net Sales) as computed number — bypasses #N/A cascade
    //    - Col L (Variance Net Sales) = 0 for reconciled stores
    await updateReconciledCellsInSummary(
      zip,
      masterWb,
      varianceResult.reconciledStores,
      gcPromoMap,
      brinkStoreData,
      lookupTableauToBrink
    );

    // Clean calcChain to prevent calculation chain repair warnings
    await cleanCalculationChain(zip);

    // Force pivot tables to refresh from source data when the file is opened.
    // This ensures Summary sheet Column C reflects the adjusted Tableau Data values.
    await enablePivotRefreshOnLoad(zip);

    // G. Generate final Franchise Sales buffer
    const outputBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    });

    const salesFileName = `Franchise Sales (Updated).xlsx`;
    const summaryText = updatedSheetsList.join(" & ");

    let maxDateSerial = 0;
    for (const row of tableauRows) {
      const s = dateToExcelSerial(row.date);
      if (s > maxDateSerial) maxDateSerial = s;
    }
    const latestDateStr = formatFilenameDate(maxDateSerial);

    return NextResponse.json({
      success: true,
      stats: {
        updatedSheet: summaryText,
        brinkRows: updatedBrinkRows,
        tableauRows: updatedTableauRows,
        totalRows: updatedBrinkRows + updatedTableauRows,
        reconciledStores: varianceResult.reconciledStoresCount,
        latestDateStr,
      },
      file: {
        fileName: salesFileName,
        data: Buffer.from(outputBuffer).toString("base64"),
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      tableauRows,
    });
  } catch (err) {
    console.error("Error updating sales files:", err);
    return NextResponse.json(
      {
        error:
          err instanceof Error
            ? err.message
            : "An unexpected error occurred while updating the sales file.",
      },
      { status: 500 }
    );
  }
}
