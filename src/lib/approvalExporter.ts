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

// Convert 1-based column index to Excel letter(s): 1→A, 26→Z, 27→AA
function colIndexToLetter(n: number): string {
  let result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

// ─── Column detection using ExcelJS (read-only — never writes back) ───────────

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
    if ('result' in (cellValue as object)) return extractCellText((cellValue as { result: ExcelJS.CellValue }).result);
    if ('formula' in (cellValue as object)) {
      const f = cellValue as { formula: string; result?: ExcelJS.CellValue };
      return f.result !== undefined ? extractCellText(f.result) : '';
    }
    if ('sharedFormula' in (cellValue as object)) {
      const sf = cellValue as { sharedFormula: string; result?: ExcelJS.CellValue };
      return sf.result !== undefined ? extractCellText(sf.result) : '';
    }
    if ('error' in (cellValue as object)) return '';
    if ('text' in (cellValue as object)) return ((cellValue as { text: string }).text ?? '').trim();
  }
  return String(cellValue).trim();
}

interface DetectedColumns {
  unitPriceCol: number; // 1-based
  totalCol: number;     // 1-based, -1 if not found
  headerRow: number;    // last header row number
}

function detectColumns(sheet: ExcelJS.Worksheet): DetectedColumns | null {
  let bestScore = 0;
  let bestRow = -1;
  let bestUnitPriceCol = -1;
  let bestTotalCol = -1;

  const getTexts = (rowNum: number): Record<number, string> => {
    const result: Record<number, string> = {};
    sheet.getRow(rowNum).eachCell({ includeEmpty: false }, (cell, col) => {
      const txt = extractCellText(cell.value);
      if (txt) result[col] = txt;
    });
    return result;
  };

  const scoreTexts = (texts: Record<number, string>): number => {
    let s = 0;
    for (const v of Object.values(texts)) {
      if (headerMatches(v, UNIT_PRICE_HEADERS)) s += 3;
      if (headerMatches(v, TOTAL_HEADERS)) s += 2;
    }
    return s;
  };

  for (let rowNum = 1; rowNum <= Math.min(60, sheet.rowCount); rowNum++) {
    const texts = getTexts(rowNum);
    const next = getTexts(rowNum + 1);
    const merged: Record<number, string> = { ...texts };
    for (const [c, v] of Object.entries(next)) {
      const col = Number(c);
      merged[col] = merged[col] ? merged[col] + ' ' + v : v;
    }

    const score = scoreTexts(merged);
    if (score < 3) continue;

    let upCol = -1;
    let totCol = -1;
    for (const [c, v] of Object.entries(merged)) {
      const col = Number(c);
      if (upCol === -1 && headerMatches(v, UNIT_PRICE_HEADERS)) upCol = col;
      if (totCol === -1 && headerMatches(v, TOTAL_HEADERS)) totCol = col;
    }

    if (upCol !== -1 && score > bestScore) {
      bestScore = score;
      bestRow = rowNum;
      bestUnitPriceCol = upCol;
      bestTotalCol = totCol;
    }
  }

  if (bestRow === -1 || bestUnitPriceCol === -1) return null;

  // Handle two-row headers
  const nextScore = scoreTexts(
    (() => {
      const r: Record<number, string> = {};
      sheet.getRow(bestRow + 1).eachCell({ includeEmpty: false }, (cell, col) => {
        const txt = extractCellText(cell.value);
        if (txt) r[col] = txt;
      });
      return r;
    })()
  );

  return {
    unitPriceCol: bestUnitPriceCol,
    totalCol: bestTotalCol,
    headerRow: nextScore >= 3 ? bestRow + 1 : bestRow,
  };
}

// ─── XML surgery helpers ──────────────────────────────────────────────────────

// Update or insert <v>value</v> in a cell element string.
// cellXml is the full <c ...>...</c> element.
function setCellValue(cellXml: string, value: number): string {
  // Remove t="s" (string type) and t="str" if present — numeric cells have no t attr
  cellXml = cellXml.replace(/\s+t="[^"]*"/, '');

  if (/<v>/.test(cellXml)) {
    // Replace existing <v>...</v>
    return cellXml.replace(/<v>[^<]*<\/v>/, `<v>${value}</v>`);
  }

  // Cell has no <v> — insert before </c>
  // Handle self-closing <c ... />
  if (/\/>$/.test(cellXml.trimEnd())) {
    return cellXml.trimEnd().replace(/\/>$/, `><v>${value}</v></c>`);
  }

  return cellXml.replace(/<\/c>/, `<v>${value}</v></c>`);
}

// Clear <v>...</v> from a cell, leaving the cell element intact (for formula cells
// we don't touch; for non-formula cells we remove the stale value).
function clearCellValue(cellXml: string): string {
  return cellXml.replace(/<v>[^<]*<\/v>/, '');
}

// Apply value mutations to xl/worksheets/sheet1.xml string.
// mutations: Map<rowIndex, Map<colLetter, number | null>>
//   number → set that value
//   null   → clear that value (leave cell structure, remove <v>)
function applyMutationsToSheetXml(
  xml: string,
  mutations: Map<number, Map<string, number | null>>
): string {
  // Process row by row using regex that captures each <row ...>...</row> block
  return xml.replace(
    /(<row\b[^>]*\br="(\d+)"[^>]*>)([\s\S]*?)(<\/row>)/g,
    (fullMatch, openTag: string, rowNumStr: string, rowBody: string, closeTag: string) => {
      const rowNum = parseInt(rowNumStr, 10);
      const colMutations = mutations.get(rowNum);
      if (!colMutations) return fullMatch; // untouched row

      let newBody = rowBody;

      for (const [colLetter, value] of colMutations) {
        const cellRef = `${colLetter}${rowNum}`;
        const esc = escapeRegex(cellRef);
        // Self-closing first, then non-self-closing with negative lookahead to avoid greedy overrun.
        // [^>]* in the opening tag cannot cross the > boundary, so r="XN" is guaranteed to be
        // in the <c ...> opening tag only — never in formula/value child text.
        const cellPattern = new RegExp(
          `<c\\b[^>]*\\br="${esc}"[^>]*\\/>`
          + `|<c\\b[^>]*\\br="${esc}"[^>]*>(?:(?!<\\/c>)[\\s\\S])*<\\/c>`,
          'g'
        );

        const existingCell = cellPattern.exec(newBody);

        if (existingCell) {
          const original = existingCell[0];
          let updated: string;
          if (value === null) {
            updated = clearCellValue(original);
          } else {
            updated = setCellValue(original, value);
          }
          newBody = newBody.slice(0, existingCell.index) + updated + newBody.slice(existingCell.index + original.length);
        } else if (value !== null) {
          // Cell doesn't exist in this row yet — append before </row>
          // We'll handle this by inserting after the last existing cell or at start
          newBody = newBody + `<c r="${cellRef}"><v>${value}</v></c>`;
        }
      }

      return openTag + newBody + closeTag;
    }
  );
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ─── Main export function — JSZip surgical XML edit, no ExcelJS write path ───

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

  // Step 1: Download original file from storage
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

  // Step 2: Use ExcelJS read-only to detect which column is "unit price"
  // We never call writeBuffer() on this workbook — it's purely for column detection.
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

  const upColLetter = colIndexToLetter(cols.unitPriceCol);
  const totColLetter = cols.totalCol !== -1 ? colIndexToLetter(cols.totalCol) : null;

  // Step 3: Build mutation map — row → col → value
  // Only touch rows we know about from the DB.
  const allBoqRowIndexes = new Set<number>(
    items.filter(i => i.row_index != null && i.row_index > 0).map(i => i.row_index)
  );

  const pricedRowMap = new Map<number, BOQItem>();
  for (const item of pricedItems) pricedRowMap.set(item.row_index, item);

  const mutations = new Map<number, Map<string, number | null>>();

  for (const rowIdx of allBoqRowIndexes) {
    const dbItem = pricedRowMap.get(rowIdx);
    const colMap = new Map<string, number | null>();

    if (dbItem) {
      // Inject unit rate
      colMap.set(upColLetter, dbItem.unit_rate!);
      // Inject total if we have a total column
      if (totColLetter) {
        const total = dbItem.total_price
          ?? Math.round((dbItem.quantity ?? 0) * (dbItem.unit_rate ?? 0) * 100) / 100;
        colMap.set(totColLetter, total);
      }
    } else {
      // Unpriced BOQ row — clear any stale values
      colMap.set(upColLetter, null);
      if (totColLetter) colMap.set(totColLetter, null);
    }

    mutations.set(rowIdx, colMap);
  }

  // Step 4: Open original xlsx as ZIP and surgically edit sheet1.xml only.
  // All other files (workbook.xml, styles.xml, shared strings, rels, etc.) pass through byte-for-byte.
  const zip = await JSZip.loadAsync(buffer);

  // Find the first sheet file (usually xl/worksheets/sheet1.xml)
  const sheetKeys = Object.keys(zip.files).filter(
    k => k.match(/^xl\/worksheets\/sheet\d+\.xml$/)
  ).sort();

  if (sheetKeys.length === 0) {
    return {
      success: false,
      injected: 0,
      total: items.length,
      variance: 0,
      unmatched: [],
      error: 'Could not locate sheet XML inside the xlsx file.',
    };
  }

  const sheetKey = sheetKeys[0];
  let sheetXml = await zip.file(sheetKey)!.async('string');

  // Apply all mutations
  sheetXml = applyMutationsToSheetXml(sheetXml, mutations);

  zip.file(sheetKey, sheetXml);

  // Remove calcChain.xml — it references stale calculation order and Excel regenerates it.
  zip.remove('xl/calcChain.xml');

  // Step 5: Generate output ZIP — all untouched files preserved byte-for-byte
  const outBuffer = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });

  triggerDownload(outBuffer, `${boqFile.name.replace(/\.xlsx?$/i, '')}_priced.xlsx`);

  await supabase
    .from('boq_files')
    .update({ export_variance_pct: 0 })
    .eq('id', boqFile.id);

  const injected = pricedItems.filter(i => i.row_index > cols.headerRow).length;
  const unmatched = pricedItems
    .filter(i => i.row_index <= cols.headerRow)
    .map(i => i.item_no || i.description.slice(0, 30));

  return {
    success: true,
    injected,
    total: items.length,
    variance: 0,
    unmatched,
  };
}
