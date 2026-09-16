import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import JSZip from "jszip";

interface TableauRow {
  date: string;
  modNo: number;
  store: string;
  gross: number;
  net: number;
  franchisee: string;
}

function escapeXml(unsafe: unknown): string {
  return String(unsafe ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
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

async function cleanCalculationChain(zip: JSZip): Promise<void> {
  zip.remove("xl/calcChain.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (relsFile) {
    let relsXml = await relsFile.async("text");
    relsXml = relsXml.replace(
      /<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/g,
      ""
    );
    zip.file("xl/_rels/workbook.xml.rels", relsXml);
  }
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ctXml = await ctFile.async("text");
    ctXml = ctXml.replace(
      /<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g,
      ""
    );
    zip.file("[Content_Types].xml", ctXml);
  }
}

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
      xml = xml.replace(
        "<pivotCacheDefinition ",
        '<pivotCacheDefinition refreshOnLoad="1" '
      );
    } else {
      xml = xml.replace(/refreshOnLoad="[^"]*"/, 'refreshOnLoad="1"');
    }
    zip.file(pcFile, xml);
  }
}

async function processRoyaltyWorkbook(
  royaltyBuf: Buffer,
  validTableauRows: TableauRow[]
): Promise<{ buffer: Buffer; fileName: string; latestDateStr: string }> {
  const zip = await JSZip.loadAsync(royaltyBuf);

  const s4File = zip.file("xl/worksheets/sheet4.xml");
  if (!s4File) {
    throw new Error(
      "Could not find sheet4.xml (Calculations) in uploaded Royalty Calculator workbook."
    );
  }
  const s4Xml = await s4File.async("text");

  const row5StartIdx = s4Xml.indexOf('<row r="5"');
  const sheetDataEndIdx = s4Xml.indexOf("</sheetData>");
  if (row5StartIdx === -1 || sheetDataEndIdx === -1) {
    throw new Error(
      "Could not locate data section in Calculations (sheet4.xml) of uploaded file."
    );
  }

  const headerXml = s4Xml.slice(0, row5StartIdx);
  const footerXml = s4Xml.slice(sheetDataEndIdx);

  // Find max date serial for Summary!B3 and Calculations side table
  let maxDateSerial = 0;
  for (const row of validTableauRows) {
    const s = dateToExcelSerial(row.date);
    if (s > maxDateSerial) maxDateSerial = s;
  }
  const startDateSerial = maxDateSerial > 6 ? maxDateSerial - 6 : maxDateSerial;

  let newRowsXml = "";
  validTableauRows.forEach((row, idx) => {
    const rNum = 5 + idx;
    const dateSerial = dateToExcelSerial(row.date);

    // Side columns for rows 5..14
    let sideColsXml = "";
    if (rNum === 5) {
      sideColsXml = `<c r="O5" s="46"><v>${startDateSerial}</v></c><c r="P5"><f>COUNTIF(B:B,O5)</f></c>`;
    } else if (rNum >= 6 && rNum <= 11) {
      sideColsXml = `<c r="O${rNum}" s="7"><f>O${rNum - 1}+1</f></c><c r="P${rNum}"><f>COUNTIF(B:B,O${rNum})</f></c>`;
    } else if (rNum === 12) {
      sideColsXml = `<c r="R12" s="11"/>`;
    } else if (rNum === 13) {
      sideColsXml = `<c r="O13" s="7"/>`;
    } else if (rNum === 14) {
      sideColsXml = `<c r="O14" s="7"/>`;
    }

    newRowsXml += `<row r="${rNum}" spans="1:18" x14ac:dyDescent="0.25">`;
    newRowsXml += `<c r="A${rNum}" s="143" t="str"><f>IF(Table11[[#This Row],[Date]]=0,"",IF(Table11[[#This Row],[Date]]&lt;(Summary!$B$3-6),1,""))</f><v/></c>`;
    newRowsXml += `<c r="B${rNum}" s="127"><v>${dateSerial}</v></c>`;
    newRowsXml += `<c r="C${rNum}" s="16"><v>${row.modNo}</v></c>`;
    newRowsXml += `<c r="D${rNum}" s="16" t="inlineStr"><is><t>${escapeXml(row.store)}</t></is></c>`;
    newRowsXml += `<c r="E${rNum}" s="17"><v>${row.gross}</v></c>`;
    newRowsXml += `<c r="F${rNum}" s="17"><v>${row.net}</v></c>`;
    newRowsXml += `<c r="G${rNum}" s="10" t="str"><f>VLOOKUP(Table11[[#This Row],[Store]],&apos;Rates from Tracker&apos;!A:B,2,0)</f></c>`;
    newRowsXml += `<c r="H${rNum}" t="str"><f>VLOOKUP(D${rNum},&apos;Rates from Tracker&apos;!A:C,3,0)</f></c>`;
    newRowsXml += `<c r="I${rNum}" t="str"><f>VLOOKUP(Table11[[#This Row],[Franchisee]],&apos;Rates from Tracker&apos;!C:D,2,0)</f></c>`;
    newRowsXml += `<c r="J${rNum}" s="10"><f>ROUND(IF(Table11[[#This Row],[Location]]=&quot;Belmont Market&quot;,Table11[[#This Row],[Column2]]*0.01,IF(Table11[[#This Row],[Location]]=&quot;Country Club Centre&quot;,Table11[[#This Row],[Column2]]*0.01,IF(Table11[[#This Row],[Location]]=&quot;Naval Air Station Norfolk&quot;,Table11[[#This Row],[Column2]]*0.005,F${rNum}*0.02))),2)</f></c>`;
    newRowsXml += `<c r="K${rNum}" s="12"><f>IF(Table11[[#This Row],[Location]]=&quot;Creekwalk&quot;,0.02,IF(Table11[[#This Row],[Location]]=&quot;Country Club Centre&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Lenexa&quot;,0.02,IF(Table11[[#This Row],[Location]]=&quot;Belmont Market&quot;,0.02,IF(Table11[[#This Row],[Location]]=&quot;Naval Air Station Norfolk&quot;,0.02,IF(Table11[[#This Row],[Location]]=&quot;Madison&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Mobile&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Decatur&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Woods Cross&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Park City&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Riverdale&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Riverton&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Layton&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Bend&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Bend&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Oakway&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;North Medford&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;South Eugene&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;American Fork&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Sugar House&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Bend North&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;South Medford&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Saratoga Springs&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Fort Union&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;South Jordan&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Grants Pass&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Pocatello&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Spanish Fork&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Downtown Boise&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Boise Town Square&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Ammon&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Roseburg&quot;,0.01,IF(Table11[[#This Row],[Location]]=&quot;Springfield-Oregon&quot;,0.01,0.05)))))))))))))))))))))))))))))))))</f></c>`;
    newRowsXml += `<c r="L${rNum}" s="11"><f>ROUND(F${rNum}*K${rNum},2)</f></c>`;
    newRowsXml += `<c r="M${rNum}" s="10"><f>J${rNum}+L${rNum}</f></c>`;
    newRowsXml += sideColsXml;
    newRowsXml += `</row>`;
  });

  const lastRow = 4 + validTableauRows.length;
  let updatedS4Xml = headerXml + newRowsXml + footerXml;
  updatedS4Xml = updatedS4Xml.replace(
    /<dimension ref="[^"]*"/,
    `<dimension ref="A1:R${lastRow}"`
  );
  zip.file("xl/worksheets/sheet4.xml", updatedS4Xml);

  // Update table3.xml (Table11)
  const t3File = zip.file("xl/tables/table3.xml");
  if (t3File) {
    let t3Xml = await t3File.async("text");
    t3Xml = t3Xml.replace(/ref="B4:M\d+"/, `ref="B4:M${lastRow}"`);
    t3Xml = t3Xml.replace(
      /<autoFilter ref="B4:M\d+"/,
      `<autoFilter ref="B4:M${lastRow}"`
    );
    t3Xml = t3Xml.replace(
      /<sortState ref="B5:M\d+"/,
      `<sortState ref="B5:M${lastRow}"`
    );
    t3Xml = t3Xml.replace(
      /<sortCondition ref="B4:B\d+"/,
      `<sortCondition ref="B4:B${lastRow}"`
    );
    zip.file("xl/tables/table3.xml", t3Xml);
  }

  // Update Summary!B3 if maxDateSerial > 0
  if (maxDateSerial > 0) {
    const s5File = zip.file("xl/worksheets/sheet5.xml");
    if (s5File) {
      let s5Xml = await s5File.async("text");
      s5Xml = s5Xml.replace(
        /(<c r="B3"[^>]*><v>)[^<]*(<\/v><\/c>)/,
        `$1${maxDateSerial}$2`
      );
      zip.file("xl/worksheets/sheet5.xml", s5Xml);
    }
  }

  // Clean calcChain
  await cleanCalculationChain(zip);

  // Enable refresh on pivot caches
  await enablePivotRefreshOnLoad(zip);

  const latestDateStr = formatFilenameDate(maxDateSerial);
  const fileName = latestDateStr
    ? `Franchise Royalty Calculator ${latestDateStr}.xlsx`
    : `Franchise Royalty Calculator.xlsx`;

  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  return { buffer, fileName, latestDateStr };
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const royaltyFile = formData.get("royaltyFile") as File | null;
    const tableauRowsJson = formData.get("tableauRows") as string | null;

    if (!royaltyFile) {
      return NextResponse.json(
        { error: "Please upload the Franchise Royalty Calculator file." },
        { status: 400 }
      );
    }

    if (!tableauRowsJson) {
      return NextResponse.json(
        {
          error:
            "No Tableau sales data found. Please complete Step 1 (Franchise Sales) first.",
        },
        { status: 400 }
      );
    }

    let tableauRows: TableauRow[] = [];
    try {
      tableauRows = JSON.parse(tableauRowsJson);
    } catch {
      return NextResponse.json(
        { error: "Invalid Tableau rows data provided." },
        { status: 400 }
      );
    }

    if (!Array.isArray(tableauRows) || tableauRows.length === 0) {
      return NextResponse.json(
        {
          error:
            "Tableau sales data contains no records. Please complete Step 1 first.",
        },
        { status: 400 }
      );
    }

    const royaltyBuf = Buffer.from(await royaltyFile.arrayBuffer());
    const res = await processRoyaltyWorkbook(royaltyBuf, tableauRows);

    return NextResponse.json({
      success: true,
      file: {
        fileName: res.fileName,
        data: Buffer.from(res.buffer).toString("base64"),
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      stats: {
        recordsCount: tableauRows.length,
        latestDateStr: res.latestDateStr,
      },
    });
  } catch (err) {
    console.error("Error processing Franchise Royalty Calculator:", err);
    return NextResponse.json(
      {
        error:
          err instanceof Error
            ? err.message
            : "An unexpected error occurred while processing the Royalty Calculator.",
      },
      { status: 500 }
    );
  }
}
