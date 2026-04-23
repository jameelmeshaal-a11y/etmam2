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

// ─── Column detection using ExcelJS (read-only — never writes back) ───────────
// Scans the first 20 rows of sheet[0] looking for any cell containing "سعر الوحدة".
// Returns the column letter (e.g. "G") or null if not found.

export async function findUnitPriceColumn(templateBuffer: ArrayBuffer): Promise<string | null> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) return null;

  const HEADERS = ['سعر الوحدة', 'سعر الوحده', 'unit price', 'unit_price', 'unitprice'];

  for (let rowNum = 1; rowNum <= Math.min(20, sheet.rowCount); rowNum++) {
    const row = sheet.getRow(rowNum);
    for (let col = 1; col <= sheet.columnCount; col++) {
      const cell = row.getCell(col);
      let text = '';
      const v = cell.value;
      if (typeof v === 'string') {
        text = v.trim();
      } else if (v && typeof v === 'object' && 'richText' in (v as object)) {
        text = (v as { richText: { text: string }[] }).richText.map(r => r.text ?? '').join('').trim();
      }
      if (!text) continue;
      const lower = text.toLowerCase();
      if (HEADERS.some(h => lower.includes(h.toLowerCase()))) {
        return sheet.getColumn(col).letter;
      }
    }
  }
  return null;
}

// ─── JSZip XML surgery — injects prices without rebuilding the workbook ───────
// Patches xl/worksheets/sheet1.xml directly and removes calcChain.xml so Excel
// recalculates formulas cleanly on open. Never touches any other part of the ZIP.

export async function injectPricesIntoXlsx(
  templateBuffer: ArrayBuffer,
  prices: Record<number, number>,
  unitPriceCol: string
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetXml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
  let patchedXml = sheetXml;

  for (const [rowNum, price] of Object.entries(prices)) {
    const cellRef = `${unitPriceCol}${rowNum}`;

    // Try to update an existing cell element
    const cellPattern = new RegExp(
      `(<c\\s[^>]*\\br="${cellRef}"[^>]*>)(<[^/].*?<\\/c>|<\\/c>)`,
      "s"
    );

    if (cellPattern.test(patchedXml)) {
      patchedXml = patchedXml.replace(cellPattern, (_, openTag) => {
        const cleanTag = openTag.replace(/\s*t="s"/, "");
        return `${cleanTag}<v>${price}</v></c>`;
      });
    } else {
      // Cell doesn't exist yet — append it inside the row element
      const rowPattern = new RegExp(
        `(<row[^>]*\\br="${rowNum}"[^>]*>)(.*?)(<\\/row>)`,
        "s"
      );
      patchedXml = patchedXml.replace(rowPattern, (_, rowOpen, rowContent, rowClose) => {
        const newCell = `<c r="${cellRef}"><v>${price}</v></c>`;
        return `${rowOpen}${rowContent}${newCell}${rowClose}`;
      });
    }
  }

  zip.file("xl/worksheets/sheet1.xml", patchedXml);
  zip.remove("xl/calcChain.xml");

  return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
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

  // Detect the unit price column letter from the original file
  const unitPriceCol = await findUnitPriceColumn(buffer);
  if (!unitPriceCol) {
    return {
      success: false, injected: 0, total: items.length, variance: 0, unmatched: [],
      error: 'تعذّر تحديد عمود سعر الوحدة في الملف. تحقق من رؤوس الأعمدة.',
    };
  }

  // Build row → price map (row_index is the 1-based Excel row number)
  const prices: Record<number, number> = {};
  for (const item of pricedItems) {
    prices[item.row_index] = item.unit_rate!;
  }

  // Inject prices via XML surgery and trigger download
  const outBuffer = await injectPricesIntoXlsx(buffer, prices, unitPriceCol);
  triggerDownload(outBuffer, `${boqFile.name.replace(/\.xlsx?$/i, '')}_priced.xlsx`);

  await supabase.from('boq_files').update({ export_variance_pct: 0 }).eq('id', boqFile.id);

  return { success: true, injected: pricedItems.length, total: items.length, variance: 0, unmatched: [] };
}
