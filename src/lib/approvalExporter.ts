import ExcelJS from 'exceljs';
import { supabase } from './supabase';
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

// Returns true if the cell contains a formula (SUM, SUBTOTAL, etc.)
function hasFormula(cellValue: ExcelJS.CellValue): boolean {
  if (!cellValue || typeof cellValue !== 'object') return false;
  const cv = cellValue as Record<string, unknown>;
  return typeof cv['formula'] === 'string'
    || typeof cv['sharedFormula'] === 'string';
}

// ─── Column detection ──────────────────────────────────────────────────────────

interface DetectedColumns {
  unitPriceCol: number;
  totalCol: number;
  headerRow: number;
  qtyCol: number;
}

const QTY_HEADERS = ['الكمية', 'كمية', 'quantity', 'qty'];

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
  let bestQtyCol = -1;

  // Scan up to 60 rows to handle files with large title sections before the table
  for (let rowNum = 1; rowNum <= Math.min(60, sheet.rowCount); rowNum++) {
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
    let qtyCol = -1;

    for (const [colStr, val] of Object.entries(merged)) {
      const colNum = Number(colStr);
      if (unitPriceCol === -1 && headerMatches(val, UNIT_PRICE_HEADERS)) unitPriceCol = colNum;
      if (totalCol === -1 && headerMatches(val, TOTAL_HEADERS)) totalCol = colNum;
      if (qtyCol === -1 && headerMatches(val, QTY_HEADERS)) qtyCol = colNum;
    }

    if (unitPriceCol !== -1 && score > bestScore) {
      bestScore = score;
      bestRow = rowNum;
      bestUnitPriceCol = unitPriceCol;
      bestTotalCol = totalCol;
      bestQtyCol = qtyCol;
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
    qtyCol: bestQtyCol,
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

  // Build set of ALL known BOQ row indexes (priced or unpriced, excluding descriptive-only)
  // Only these rows may have their price cells cleared — structural rows are never touched
  const allBoqRowIndexes = new Set<number>(
    items
      .filter(i => i.row_index != null && i.row_index > 0)
      .map(i => i.row_index)
  );

  // Compute grand total from DB
  const dbGrandTotal = pricedItems.reduce(
    (sum, i) => sum + (i.total_price ?? (i.quantity ?? 0) * (i.unit_rate ?? 0)),
    0
  );
  const dbGrandTotalRounded = Math.round(dbGrandTotal * 100) / 100;

  // Step 4: Find grand total row — last hard-coded numeric cell in totalCol whose
  // value is plausibly a grand total (>= 10% of DB total). Formula grand totals
  // will recalculate automatically and don't need to be overwritten.
  let grandTotalRowNum = -1;
  if (cols.totalCol !== -1 && dbGrandTotal > 0) {
    sheet.eachRow((row, rowNum) => {
      if (rowNum <= cols.headerRow) return;
      const totCell = row.getCell(cols.totalCol);
      if (!totCell.value || hasFormula(totCell.value)) return;
      const numVal = typeof totCell.value === 'number'
        ? totCell.value
        : typeof totCell.value === 'object' && 'result' in (totCell.value as object)
          ? (totCell.value as { result: ExcelJS.CellValue }).result as number
          : null;
      if (typeof numVal === 'number' && numVal > 0 && numVal >= dbGrandTotal * 0.1) {
        grandTotalRowNum = rowNum;
      }
    });
  }

  // Step 5: Inject prices — conservative approach:
  //   - Formula cells: always leave untouched (Excel recalculates on open)
  //   - Known BOQ rows with DB price: inject unit_rate + total_price
  //   - Known BOQ rows without DB price: clear price cells (previously unpriced)
  //   - ALL other rows (structural, subtotals, section headers): leave completely untouched
  //   - Grand total row with hard-coded number: overwrite with DB grand total
  let injected = 0;
  const unmatched: string[] = [];

  sheet.eachRow((row, rowNum) => {
    if (rowNum <= cols.headerRow) return;

    const dbItem = pricedRowMap.get(rowNum);
    const isKnownBoqRow = allBoqRowIndexes.has(rowNum);
    const isGrandTotalRow = rowNum === grandTotalRowNum;

    // Only touch rows we actually need to modify — never commit untouched rows
    let rowModified = false;

    // Grand total row: overwrite with DB grand total
    if (isGrandTotalRow && cols.totalCol !== -1) {
      const totCell = row.getCell(cols.totalCol);
      if (!hasFormula(totCell.value)) {
        totCell.value = dbGrandTotalRounded;
        rowModified = true;
      }
    } else {
      // ── Unit price column ──────────────────────────────────────────────────
      const upCell = row.getCell(cols.unitPriceCol);
      if (!hasFormula(upCell.value)) {
        if (dbItem) {
          upCell.value = dbItem.unit_rate;
          rowModified = true;
        } else if (isKnownBoqRow && upCell.value != null) {
          upCell.value = null;
          rowModified = true;
        }
      }

      // ── Total column ───────────────────────────────────────────────────────
      if (cols.totalCol !== -1) {
        const totCell = row.getCell(cols.totalCol);
        if (!hasFormula(totCell.value)) {
          if (dbItem) {
            const total = dbItem.total_price
              ?? Math.round((dbItem.quantity ?? 0) * (dbItem.unit_rate ?? 0) * 100) / 100;
            totCell.value = total;
            rowModified = true;
          } else if (isKnownBoqRow && totCell.value != null) {
            totCell.value = null;
            rowModified = true;
          }
        }
      }
    }

    if (rowModified) row.commit();
    if (dbItem) injected++;
  });

  // Collect items whose row_index fell at or before the header (shouldn't happen normally)
  for (const item of pricedItems) {
    if (item.row_index <= cols.headerRow) {
      unmatched.push(item.item_no || item.description.slice(0, 30));
    }
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
