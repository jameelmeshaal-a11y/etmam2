// ⛔ LOCKED FILE — DO NOT MODIFY WITHOUT EXPLICIT USER PERMISSION
// Last approved state: 2026-04-23
// Any change to this file requires user to say: "افتح ملف approvalExporter.ts"
import ExcelJS from 'exceljs';
// @ts-expect-error — jszip ships no bundled types but is present as ExcelJS dependency
import JSZip from 'jszip';
import { supabase } from './supabase';
import type { BOQItem, BOQFile, ExportResult } from '../types';

// ─── Export unpriced items for rate library upload ───────────────────────────
// Uses ExcelJS to BUILD a new file from scratch — no round-trip corruption risk.

export async function exportUnpricedItemsForLibrary(
  boqFileName: string,
  items: BOQItem[]
): Promise<void> {
  const unpriced = items.filter(
    i => i.status !== 'descriptive'
      && (i.quantity ?? 0) > 0
      && (i.unit_rate == null || i.unit_rate === 0)
  );

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('البنود غير المسعرة');

  sheet.views = [{ rightToLeft: true }];

  sheet.columns = [
    { header: 'رقم البند', key: 'item_no', width: 18 },
    { header: 'وصف البند', key: 'description', width: 55 },
    { header: 'الوحدة', key: 'unit', width: 12 },
    { header: 'الكمية', key: 'quantity', width: 12 },
    { header: 'سعر الوحدة المقترح', key: 'rate_target', width: 22 },
    { header: 'الحد الأدنى', key: 'rate_min', width: 15 },
    { header: 'الحد الأقصى', key: 'rate_max', width: 15 },
    { header: 'التصنيف', key: 'category', width: 18 },
    { header: 'الاسم المعياري', key: 'standard_name_ar', width: 50 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 28;

  unpriced.forEach(item => {
    const row = sheet.addRow({
      item_no: item.item_no || '',
      description: item.description,
      unit: item.unit || '',
      quantity: item.quantity ?? 0,
      rate_target: '',
      rate_min: '',
      rate_max: '',
      category: '',
      standard_name_ar: item.description,
    });

    ['rate_target', 'rate_min', 'rate_max', 'category'].forEach(key => {
      const cell = row.getCell(key);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    });

    row.alignment = { vertical: 'middle', wrapText: true };
    row.height = 22;
  });

  sheet.views[0].state = 'frozen';
  sheet.views[0].ySplit = 1;

  const outBuffer = await workbook.xlsx.writeBuffer();
  triggerDownload(
    outBuffer,
    `${boqFileName.replace(/\.xlsx?$/i, '')}_unpriced_for_library.xlsx`
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function triggerDownload(buffer: ArrayBuffer | Buffer, filename: string): void {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

const UP_HEADERS = ['سعر الوحدة', 'سعر الوحده', 'unit price', 'unit_price', 'unitprice'];
const BOQ_SCORE_HEADERS = [
  ['وصف البند', 'وصف', 'البيان', 'الوصف', 'description'],
  ['الكمية', 'كمية', 'qty', 'quantity'],
  UP_HEADERS,
];

function extractText(v: ExcelJS.CellValue): string {
  if (!v) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'object' && 'richText' in (v as object))
    return (v as { richText: { text: string }[] }).richText.map(r => r.text ?? '').join('').trim();
  return '';
}

function sheetScore(sheet: ExcelJS.Worksheet): { score: number; colLetter: string | null } {
  let best = { score: 0, colLetter: null as string | null };
  for (let r = 1; r <= Math.min(30, sheet.rowCount); r++) {
    let score = 0;
    let colLetter: string | null = null;
    for (let pass = 0; pass < 2; pass++) {
      sheet.getRow(r + pass).eachCell({ includeEmpty: false }, (cell, col) => {
        const lower = extractText(cell.value).toLowerCase();
        if (!lower) return;
        for (const group of BOQ_SCORE_HEADERS) {
          if (group.some(h => lower.includes(h.toLowerCase()))) { score++; break; }
        }
        if (!colLetter && UP_HEADERS.some(h => lower.includes(h.toLowerCase())))
          colLetter = sheet.getColumn(col).letter;
      });
    }
    if (score > best.score && colLetter) best = { score, colLetter };
  }
  return best;
}

// ─── Column detection using ExcelJS (read-only — never writes back) ───────────
// Picks the sheet with the strongest BOQ header match (same logic as the parser),
// so the column letter and sheet index are consistent with how row_index was stored.
// Returns { colLetter, sheetXmlPath } or null if the header is not found.

export async function findUnitPriceColumn(
  templateBuffer: ArrayBuffer
): Promise<{ colLetter: string; sheetXmlPath: string } | null> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  // Identify which sheet index has the best BOQ score
  let bestScore = 0;
  let bestCol: string | null = null;
  let bestSheetIndex = -1;

  for (let si = 0; si < workbook.worksheets.length; si++) {
    const { score, colLetter } = sheetScore(workbook.worksheets[si]);
    if (score > bestScore && colLetter) {
      bestScore = score;
      bestCol = colLetter;
      bestSheetIndex = si;
    }
  }

  if (!bestCol || bestSheetIndex < 0) return null;

  // Map ExcelJS sheet index → ZIP path via workbook.xml.rels
  // @ts-expect-error — jszip ships no bundled types
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetXmlPath = await resolveSheetZipPath(zip, bestSheetIndex);

  return { colLetter: bestCol, sheetXmlPath };
}

// Resolve workbook sheet index (0-based) → ZIP-internal path (e.g. "xl/worksheets/sheet2.xml")
// by reading xl/workbook.xml and xl/_rels/workbook.xml.rels.
async function resolveSheetZipPath(
  // @ts-expect-error — jszip ships no bundled types
  zip: unknown,
  sheetIndex: number
): Promise<string> {
  try {
    // @ts-expect-error — dynamic JSZip API
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const relMap: Record<string, string> = {};
    const relPattern = /<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"/g;
    let m: RegExpExecArray | null;
    while ((m = relPattern.exec(relsXml)) !== null) relMap[m[1]] = m[2];

    // @ts-expect-error — dynamic JSZip API
    const wbXml = await zip.file("xl/workbook.xml").async("string");
    const sheetPattern = /<sheet\b[^>]+r:id="([^"]+)"[^>]*/g;
    let idx = 0;
    while ((m = sheetPattern.exec(wbXml)) !== null) {
      if (idx === sheetIndex) {
        const target = relMap[m[1]];
        if (target) return target.startsWith('xl/') ? target : `xl/${target}`;
      }
      idx++;
    }
  } catch (_) { /* fallback below */ }
  return "xl/worksheets/sheet1.xml";
}

// ─── JSZip XML surgery — injects prices without rebuilding the workbook ───────
// Patches the target sheet XML directly. Never rebuilds the workbook.
// Key correctness rules:
//   1. Replace the ENTIRE <c .../> or <c ...>...</c> element with a clean numeric
//      cell — this correctly handles shared-string cells (t="s"), formula cells,
//      self-closing cells, and any other variant.
//   2. Do NOT force compression — let JSZip preserve each file's original method
//      so Excel's ZIP reader sees an identical central directory structure.
//   3. Remove calcChain.xml and set fullCalcOnLoad in workbook.xml so Excel
//      recalculates all formula totals the moment the file is opened.

export async function injectPricesIntoXlsx(
  templateBuffer: ArrayBuffer,
  prices: Record<number, number>,
  unitPriceCol: string,
  sheetXmlPath = "xl/worksheets/sheet1.xml"
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);

  // @ts-expect-error — dynamic JSZip API
  const sheetFile = zip.file(sheetXmlPath);
  if (!sheetFile) throw new Error(`Sheet not found in ZIP: ${sheetXmlPath}`);
  const sheetXml: string = await sheetFile.async("string");
  let patchedXml = sheetXml;

  for (const [rowNum, price] of Object.entries(prices)) {
    const cellRef = `${unitPriceCol}${rowNum}`;
    // Match ANY form of the cell: self-closing or with children, any attributes
    // Replace the entire element with a minimal numeric cell — no t="s", no formula
    const cellPattern = new RegExp(
      `<c\\b[^>]*\\br="${cellRef}"[^>]*(\\/>|>[\\s\\S]*?<\\/c>)`,
      "g"
    );

    if (cellPattern.test(patchedXml)) {
      // Reset lastIndex after .test()
      patchedXml = patchedXml.replace(
        new RegExp(`<c\\b[^>]*\\br="${cellRef}"[^>]*(\\/>|>[\\s\\S]*?<\\/c>)`, "g"),
        `<c r="${cellRef}"><v>${price}</v></c>`
      );
    } else {
      // Cell doesn't exist in this row — insert it inside the row element
      const rowPattern = new RegExp(
        `(<row[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(<\\/row>)`,
        "g"
      );
      patchedXml = patchedXml.replace(rowPattern, (_, rowOpen, rowContent, rowClose) => {
        return `${rowOpen}${rowContent}<c r="${cellRef}"><v>${price}</v></c>${rowClose}`;
      });
    }
  }

  // @ts-expect-error — dynamic JSZip API
  zip.file(sheetXmlPath, patchedXml);

  // Remove calcChain so Excel doesn't try to validate stale calculation order
  // @ts-expect-error — dynamic JSZip API
  zip.remove("xl/calcChain.xml");

  // Tell Excel to recalculate all formulas on open (so totals reflect injected prices)
  try {
    // @ts-expect-error — dynamic JSZip API
    const wbFile = zip.file("xl/workbook.xml");
    if (wbFile) {
      const wbXml: string = await wbFile.async("string");
      let updatedWb = wbXml;
      if (/<calcPr\b/.test(updatedWb)) {
        updatedWb = updatedWb.replace(/<calcPr\b[^>]*\/>/,
          '<calcPr calcCompleted="0" calcMode="auto" fullCalcOnLoad="1"/>');
      } else {
        updatedWb = updatedWb.replace('</workbook>', '<calcPr calcCompleted="0" calcMode="auto" fullCalcOnLoad="1"/></workbook>');
      }
      // @ts-expect-error — dynamic JSZip API
      zip.file("xl/workbook.xml", updatedWb);
    }
  } catch (_) { /* non-fatal: file still opens, just needs manual recalc */ }

  // IMPORTANT: do NOT pass compression option — preserves each file's original
  // compression method (STORE vs DEFLATE), preventing ZIP central-directory corruption
  return await zip.generateAsync({ type: "arraybuffer" });
}

// ─── exportBOQ — public API used by BOQTable ─────────────────────────────────

export async function exportBOQ(boqFile: BOQFile, items: BOQItem[]): Promise<ExportResult> {
  const pricedItems = items.filter(
    i => i.unit_rate != null && i.unit_rate > 0
      && i.status !== 'descriptive'
      && (i.quantity ?? 0) > 0
      && i.row_index != null && i.row_index > 0
  );

  if (pricedItems.length === 0) {
    return {
      success: false, injected: 0, total: items.length, variance: 0, unmatched: [],
      error: 'لا توجد بنود مسعّرة للتصدير. يرجى تسعير البنود أولاً.',
    };
  }

  // Download original file from Supabase Storage
  let buffer: ArrayBuffer;
  try {
    const { data, error } = await supabase.storage.from('boq-files').download(boqFile.storage_path);
    if (error || !data) throw new Error(error?.message ?? 'Failed to download file');
    buffer = await data.arrayBuffer();
  } catch (e) {
    return {
      success: false, injected: 0, total: items.length, variance: 0, unmatched: [],
      error: `تعذّر تحميل الملف الأصلي: ${(e as Error).message}`,
    };
  }

  // Detect the unit price column letter and sheet path from the original file
  const colResult = await findUnitPriceColumn(buffer);
  if (!colResult) {
    return {
      success: false, injected: 0, total: items.length, variance: 0, unmatched: [],
      error: 'تعذّر تحديد عمود سعر الوحدة في الملف. تحقق من رؤوس الأعمدة.',
    };
  }

  // Build row → price map (row_index is the 1-based Excel row number stored during parsing)
  const prices: Record<number, number> = {};
  for (const item of pricedItems) {
    prices[item.row_index] = item.unit_rate!;
  }

  // Inject prices via XML surgery on the correct sheet and trigger download
  const outBuffer = await injectPricesIntoXlsx(buffer, prices, colResult.colLetter, colResult.sheetXmlPath);
  triggerDownload(outBuffer, `${boqFile.name.replace(/\.xlsx?$/i, '')}_priced.xlsx`);

  await supabase.from('boq_files').update({ export_variance_pct: 0 }).eq('id', boqFile.id);

  return { success: true, injected: pricedItems.length, total: items.length, variance: 0, unmatched: [] };
}
