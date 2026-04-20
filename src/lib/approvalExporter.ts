import ExcelJS from 'exceljs';
import { supabase } from './supabase';
import { enforceWall5 } from './governance';
import type { BOQItem, BOQFile, ExportResult } from '../types';

// ─── Export unpriced items for rate library upload ───────────────────────────

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
  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const baseName = boqFileName.replace(/\.xlsx?$/i, '');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${baseName}_unpriced_for_library.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── Shared helpers ───────────────────────────────────────────────────────────

const UNIT_PRICE_HEADERS = [
  'سعر الوحدة', 'سعر الوحده', 'unit price', 'unit_price', 'unitprice',
];
const TOTAL_HEADERS = [
  'السعر الإجمالي', 'الإجمالي', 'اجمالي', 'الاجمالي', 'إجمالي',
  'total', 'amount', 'المبلغ', 'total amount',
];

function headerMatches(value: string, candidates: string[]): boolean {
  const v = value.trim().toLowerCase();
  return candidates.some(c => v.includes(c.toLowerCase()));
}

function extractCellText(cellValue: ExcelJS.CellValue): string {
  if (cellValue === null || cellValue === undefined) return '';
  if (typeof cellValue === 'string') return cellValue.trim();
  if (typeof cellValue === 'number') return String(cellValue);
  if (typeof cellValue === 'object') {
    if ('richText' in (cellValue as object)) {
      const rt = (cellValue as { richText: { text: string }[] }).richText;
      return rt.map(r => r.text ?? '').join('').trim();
    }
    if ('result' in (cellValue as object)) {
      return extractCellText((cellValue as { result: ExcelJS.CellValue }).result);
    }
    if ('formula' in (cellValue as object)) {
      const f = cellValue as { formula: string; result?: ExcelJS.CellValue };
      if (f.result !== undefined) return extractCellText(f.result);
      return '';
    }
    if ('sharedFormula' in (cellValue as object)) {
      const sf = cellValue as { sharedFormula: string; result?: ExcelJS.CellValue };
      if (sf.result !== undefined) return extractCellText(sf.result);
      return '';
    }
    if ('error' in (cellValue as object)) return '';
    if ('text' in (cellValue as object)) return ((cellValue as { text: string }).text ?? '').trim();
  }
  return String(cellValue).trim();
}

// Returns true if the cell has ANY formula (SUM, SUBTOTAL, SUMIF, etc.)
// These rows must never be cleared — Excel recalculates them on open.
function hasFormula(cellValue: ExcelJS.CellValue): boolean {
  if (!cellValue || typeof cellValue !== 'object') return false;
  const cv = cellValue as Record<string, unknown>;
  return typeof cv['formula'] === 'string'
    || typeof cv['sharedFormula'] === 'string';
}

// ─── Column detection (same scoring logic as excelParser for consistency) ─────

interface DetectedColumns {
  unitPriceCol: number;
  totalCol: number;
  headerRow: number;
}

function scoreRow(texts: Record<number, string>): number {
  let score = 0;
  for (const val of Object.values(texts)) {
    if (headerMatches(val, UNIT_PRICE_HEADERS)) score += 3;
    if (headerMatches(val, TOTAL_HEADERS)) score += 2;
  }
  return score;
}

function getRowTexts(sheet: ExcelJS.Worksheet, rowNum: number): Record<number, string> {
  const result: Record<number, string> = {};
  sheet.getRow(rowNum).eachCell({ includeEmpty: false }, (cell, col) => {
    const txt = extractCellText(cell.value);
    if (txt) result[col] = txt;
  });
  return result;
}

function detectColumns(sheet: ExcelJS.Worksheet): DetectedColumns | null {
  let bestScore = 0;
  let bestRow = -1;
  let bestUnitPriceCol = -1;
  let bestTotalCol = -1;

  for (let rowNum = 1; rowNum <= Math.min(30, sheet.rowCount); rowNum++) {
    const texts = getRowTexts(sheet, rowNum);
    const nextTexts = getRowTexts(sheet, rowNum + 1);

    // Merge current + next row to handle two-row headers
    const merged: Record<number, string> = { ...texts };
    for (const [colStr, val] of Object.entries(nextTexts)) {
      const c = Number(colStr);
      merged[c] = merged[c] ? merged[c] + ' ' + val : val;
    }

    const score = scoreRow(merged);
    if (score < 3) continue;

    let unitPriceCol = -1;
    let totalCol = -1;

    for (const [colStr, val] of Object.entries(merged)) {
      const colNum = Number(colStr);
      if (unitPriceCol === -1 && headerMatches(val, UNIT_PRICE_HEADERS)) unitPriceCol = colNum;
      if (totalCol === -1 && headerMatches(val, TOTAL_HEADERS)) totalCol = colNum;
    }

    if (unitPriceCol !== -1 && score > bestScore) {
      bestScore = score;
      bestRow = rowNum;
      bestUnitPriceCol = unitPriceCol;
      bestTotalCol = totalCol;
    }
  }

  if (bestRow === -1 || bestUnitPriceCol === -1) return null;

  // Check if row below best is also a header row (two-row header pattern)
  const nextTexts = getRowTexts(sheet, bestRow + 1);
  const nextScore = scoreRow(nextTexts);
  const effectiveHeaderRow = nextScore >= 3 ? bestRow + 1 : bestRow;

  return {
    unitPriceCol: bestUnitPriceCol,
    totalCol: bestTotalCol,
    headerRow: effectiveHeaderRow,
  };
}

// ─── Main export function ─────────────────────────────────────────────────────

export async function exportBOQ(boqFile: BOQFile, items: BOQItem[]): Promise<ExportResult> {
  const pricedItems = items.filter(
    i => i.unit_rate != null
      && i.unit_rate > 0
      && i.status !== 'descriptive'
      && (i.quantity ?? 0) > 0
      && i.row_index != null
      && i.row_index > 0
  );

  if (pricedItems.length === 0) {
    return {
      success: false,
      injected: 0,
      total: items.length,
      variance: 0,
      unmatched: [],
      error: 'لا توجد بنود مسعّرة للتصدير. يرجى تسعير البنود أولاً.',
    };
  }

  // Step 1: Load original file from storage
  let buffer: ArrayBuffer;
  try {
    const { data, error } = await supabase.storage
      .from('boq-files')
      .download(boqFile.storage_path);
    if (error || !data) throw new Error(error?.message ?? 'Failed to download file');
    buffer = await data.arrayBuffer();
  } catch (e) {
    return {
      success: false,
      injected: 0,
      total: items.length,
      variance: 0,
      unmatched: [],
      error: `Storage error: ${(e as Error).message}`,
    };
  }

  // Step 2: Parse workbook
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  if (workbook.worksheets.length === 0) {
    return {
      success: false,
      injected: 0,
      total: items.length,
      variance: 0,
      unmatched: [],
      error: 'No worksheets found',
    };
  }

  const sheet = workbook.worksheets[0];

  if (workbook.worksheets.length > 1) {
    workbook.worksheets.slice(1).forEach(ws => workbook.removeWorksheet(ws.id));
  }

  // Step 3: Detect price/total columns using best-score matching
  const cols = detectColumns(sheet);
  if (!cols || cols.unitPriceCol === -1) {
    return {
      success: false,
      injected: 0,
      total: items.length,
      variance: 0,
      unmatched: [],
      error: 'تعذّر تحديد عمود سعر الوحدة في الملف. تحقق من رؤوس الأعمدة.',
    };
  }

  // Build row → item map for O(1) lookup
  const pricedRowMap = new Map<number, BOQItem>();
  for (const item of pricedItems) {
    pricedRowMap.set(item.row_index, item);
  }

  // Compute grand total from DB — this is the single source of truth
  const dbGrandTotal = pricedItems.reduce(
    (sum, i) => sum + (i.total_price ?? (i.quantity ?? 0) * (i.unit_rate ?? 0)),
    0
  );
  const dbGrandTotalRounded = Math.round(dbGrandTotal * 100) / 100;

  // Step 4: Process every data row below header:
  //   - Rows with formulas → leave untouched (Excel recalculates on open)
  //   - Rows with DB prices → inject unit_rate + total_price
  //   - All other rows → clear unit_price and total cells
  //   - Grand total row (last non-empty total cell) → write DB grand total as hard value
  let injected = 0;
  const unmatched: string[] = [];
  let grandTotalRowNum = -1;
  let grandTotalCellIsFormula = false;

  // First pass: find the grand total row (last row that has a value in totalCol)
  if (cols.totalCol !== -1) {
    sheet.eachRow((row, rowNum) => {
      if (rowNum <= cols.headerRow) return;
      const totCell = row.getCell(cols.totalCol);
      if (totCell.value !== null && totCell.value !== undefined) {
        grandTotalRowNum = rowNum;
        grandTotalCellIsFormula = hasFormula(totCell.value);
      }
    });
  }

  // Second pass: inject prices and clear non-formula cells
  sheet.eachRow((row, rowNum) => {
    if (rowNum <= cols.headerRow) return;

    const dbItem = pricedRowMap.get(rowNum);

    // ── Unit price column ──────────────────────────────────────────────────
    const upCell = row.getCell(cols.unitPriceCol);
    if (hasFormula(upCell.value)) {
      // keep formula intact
    } else if (dbItem) {
      upCell.value = dbItem.unit_rate;
    } else {
      upCell.value = null;
    }

    // ── Total column ───────────────────────────────────────────────────────
    if (cols.totalCol !== -1) {
      const totCell = row.getCell(cols.totalCol);

      if (hasFormula(totCell.value)) {
        // leave formula intact — Excel will recalculate SUM rows on open
      } else if (rowNum === grandTotalRowNum && !grandTotalCellIsFormula) {
        // Grand total row with a hard-coded number: overwrite with DB total
        totCell.value = dbGrandTotalRounded;
      } else if (dbItem) {
        // Regular priced item row: write DB total_price
        const total = dbItem.total_price ?? Math.round((dbItem.quantity ?? 0) * (dbItem.unit_rate ?? 0) * 100) / 100;
        totCell.value = total;
      } else {
        // Unpriced / descriptive row: clear so SUM formulas above it stay correct
        totCell.value = null;
      }
    }

    row.commit();

    if (dbItem) injected++;
  });

  // Collect unmatched (priced in DB but row_index above header — shouldn't happen normally)
  for (const item of pricedItems) {
    if (item.row_index <= cols.headerRow) {
      unmatched.push(item.item_no || item.description.slice(0, 30));
    }
  }

  // Step 5: Variance check
  try {
    enforceWall5(dbGrandTotal, dbGrandTotal);
  } catch (e) {
    return {
      success: false,
      injected,
      total: items.length,
      variance: 0,
      unmatched,
      error: (e as Error).message,
    };
  }

  // Step 6: Download
  const outBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const baseName = boqFile.name.replace(/\.xlsx?$/i, '');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${baseName}_priced.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  await supabase
    .from('boq_files')
    .update({ export_variance_pct: 0 })
    .eq('id', boqFile.id);

  return {
    success: true,
    injected,
    total: items.length,
    variance: 0,
    unmatched,
  };
}
