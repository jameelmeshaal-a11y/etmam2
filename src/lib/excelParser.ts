import ExcelJS from 'exceljs';

export interface ParsedBOQItem {
  item_no: string;
  description: string;
  unit: string;
  quantity: number | null;
  row_index: number;
  is_descriptive: boolean;
}

export interface DetectedColumns {
  itemNoCol: number;
  descCol: number;
  unitCol: number;
  qtyCol: number;
  unitPriceCol: number;
  totalCol: number;
}

export interface ParseResult {
  items: ParsedBOQItem[];
  headerRow: number;
  columns: DetectedColumns;
}

const ITEM_NO_HEADERS = ['رقم البند', 'رقم القسم', 'رقم', 'البند', 'item no', 'item number', 'no', 'division', 'بند', 'division no'];
const DESC_HEADERS = ['وصف البند', 'وصف', 'البيان', 'الوصف', 'description', 'desc', 'item description', 'بالعربية'];
const UNIT_HEADERS = ['الوحدة', 'وحدة', 'unit'];
const QTY_HEADERS = ['الكمية', 'كمية', 'qty', 'quantity'];
const UNIT_PRICE_HEADERS = ['سعر الوحدة', 'سعر الوحده', 'unit price', 'unit_price', 'unitprice'];
const TOTAL_HEADERS = ['السعر الإجمالي', 'الإجمالي', 'اجمالي', 'الاجمالي', 'إجمالي', 'total', 'amount', 'المبلغ', 'total amount'];

function headerMatches(value: string, candidates: string[]): boolean {
  const v = value.trim().toLowerCase();
  return candidates.some(c => v.includes(c.toLowerCase()));
}

function extractCellText(cellValue: ExcelJS.CellValue): string {
  if (cellValue === null || cellValue === undefined) return '';
  if (typeof cellValue === 'string') return cellValue.trim();
  if (typeof cellValue === 'number') return String(cellValue);
  if (typeof cellValue === 'boolean') return String(cellValue);
  if (cellValue instanceof Date) return cellValue.toLocaleDateString();
  if (typeof cellValue === 'object') {
    if ('richText' in (cellValue as object)) {
      const rt = (cellValue as { richText: { text: string }[] }).richText;
      return rt.map(r => r.text ?? '').join('').trim();
    }
    if ('result' in (cellValue as object)) {
      const r = (cellValue as { result: ExcelJS.CellValue }).result;
      return extractCellText(r);
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

function extractCellNumber(cellValue: ExcelJS.CellValue): number | null {
  if (cellValue === null || cellValue === undefined) return null;
  if (typeof cellValue === 'number') return isNaN(cellValue) ? null : cellValue;
  if (typeof cellValue === 'object' && cellValue !== null) {
    if ('result' in (cellValue as object)) {
      return extractCellNumber((cellValue as { result: ExcelJS.CellValue }).result);
    }
    if ('formula' in (cellValue as object)) {
      const f = cellValue as { formula: string; result?: ExcelJS.CellValue };
      if (f.result !== undefined) return extractCellNumber(f.result);
      return null;
    }
    if ('sharedFormula' in (cellValue as object)) {
      const sf = cellValue as { sharedFormula: string; result?: ExcelJS.CellValue };
      if (sf.result !== undefined) return extractCellNumber(sf.result);
      return null;
    }
  }
  const str = extractCellText(cellValue);
  const num = parseFloat(str.replace(/,/g, ''));
  return isNaN(num) ? null : num;
}

function getRowText(sheet: ExcelJS.Worksheet, rowNum: number): Record<number, string> {
  const result: Record<number, string> = {};
  const row = sheet.getRow(rowNum);
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    const txt = extractCellText(cell.value);
    if (txt) result[col] = txt;
  });
  return result;
}

function scoreHeaderRow(cellTexts: Record<number, string>): number {
  let score = 0;
  for (const val of Object.values(cellTexts)) {
    if (headerMatches(val, DESC_HEADERS)) score += 3;
    if (headerMatches(val, ITEM_NO_HEADERS)) score += 2;
    if (headerMatches(val, UNIT_HEADERS)) score += 2;
    if (headerMatches(val, QTY_HEADERS)) score += 2;
    if (headerMatches(val, UNIT_PRICE_HEADERS)) score += 2;
    if (headerMatches(val, TOTAL_HEADERS)) score += 1;
  }
  return score;
}

function detectColumns(sheet: ExcelJS.Worksheet): { cols: DetectedColumns; headerRow: number } | null {
  let bestScore = 0;
  let bestRow = -1;
  let bestCols: DetectedColumns | null = null;

  for (let rowNum = 1; rowNum <= Math.min(30, sheet.rowCount); rowNum++) {
    const cellTexts = getRowText(sheet, rowNum);
    const nextTexts = getRowText(sheet, rowNum + 1);

    const merged: Record<number, string> = { ...cellTexts };
    for (const [col, val] of Object.entries(nextTexts)) {
      const c = Number(col);
      merged[c] = merged[c] ? merged[c] + ' ' + val : val;
    }

    const score = scoreHeaderRow(merged);
    if (score < 3) continue;

    const cols: DetectedColumns = {
      itemNoCol: -1,
      descCol: -1,
      unitCol: -1,
      qtyCol: -1,
      unitPriceCol: -1,
      totalCol: -1,
    };

    for (const [colStr, val] of Object.entries(merged)) {
      const colNum = Number(colStr);
      if (cols.itemNoCol === -1 && headerMatches(val, ITEM_NO_HEADERS)) cols.itemNoCol = colNum;
      if (cols.descCol === -1 && headerMatches(val, DESC_HEADERS)) cols.descCol = colNum;
      if (cols.unitCol === -1 && headerMatches(val, UNIT_HEADERS)) cols.unitCol = colNum;
      if (cols.qtyCol === -1 && headerMatches(val, QTY_HEADERS)) cols.qtyCol = colNum;
      if (cols.unitPriceCol === -1 && headerMatches(val, UNIT_PRICE_HEADERS)) cols.unitPriceCol = colNum;
      if (cols.totalCol === -1 && headerMatches(val, TOTAL_HEADERS)) cols.totalCol = colNum;
    }

    if (cols.descCol === -1 && cols.itemNoCol === -1) continue;

    if (score > bestScore) {
      bestScore = score;
      bestRow = rowNum;
      bestCols = cols;
    }
  }

  if (!bestCols || bestRow === -1) return null;

  const nextTexts = getRowText(sheet, bestRow + 1);
  const nextScore = scoreHeaderRow(nextTexts);
  const effectiveHeaderRow = nextScore >= 2 ? bestRow + 1 : bestRow;

  return { cols: bestCols, headerRow: effectiveHeaderRow };
}

function getCellString(row: ExcelJS.Row, colNum: number): string {
  if (colNum === -1) return '';
  return extractCellText(row.getCell(colNum).value);
}

function getCellNumber(row: ExcelJS.Row, colNum: number): number | null {
  if (colNum === -1) return null;
  return extractCellNumber(row.getCell(colNum).value);
}

function isDescriptiveRow(
  description: string,
  unit: string,
  quantity: number | null
): boolean {
  if (!description) return false;
  if (unit !== '' && quantity !== null && quantity > 0) return false;
  if (quantity !== null && quantity > 0) return false;
  const hasNumbers = /\d/.test(description);
  if (hasNumbers && unit !== '') return false;
  const sectionPatterns = [
    /^[A-Z]\s*[-–]/,
    /^(قسم|بند|فصل|section|division|part)\s/i,
    /^\d+[\.\-]\s*[^\d]/,
    /^[أ-ي]\s*[-–]/,
  ];
  for (const p of sectionPatterns) {
    if (p.test(description)) return true;
  }
  if (description.length > 5 && unit === '' && (quantity === null || quantity === 0)) {
    return true;
  }
  return false;
}

export async function parseEtimadBOQ(buffer: ArrayBuffer): Promise<ParseResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('لم يتم العثور على أي صفحة في ملف Excel');
  }

  const detected = detectColumns(sheet);
  if (!detected) {
    throw new Error('تعذر اكتشاف أعمدة BOQ في الملف. تأكد من وجود رؤوس الأعمدة المطلوبة.');
  }

  const { cols, headerRow } = detected;
  const items: ParsedBOQItem[] = [];

  sheet.eachRow((row, rowNum) => {
    if (rowNum <= headerRow) return;

    const description = getCellString(row, cols.descCol);
    const item_no = getCellString(row, cols.itemNoCol);

    if (!description && !item_no) return;

    if (headerMatches(description, DESC_HEADERS) || headerMatches(item_no, ITEM_NO_HEADERS)) {
      if (headerMatches(description, UNIT_HEADERS) || headerMatches(description, QTY_HEADERS)) return;
      const rowScore = scoreHeaderRow({ 1: description, 2: item_no });
      if (rowScore >= 3) return;
    }

    const unit = getCellString(row, cols.unitCol);
    const quantity = getCellNumber(row, cols.qtyCol);
    const is_descriptive = isDescriptiveRow(description, unit, quantity);

    items.push({
      item_no,
      description,
      unit,
      quantity: is_descriptive ? null : quantity,
      row_index: rowNum,
      is_descriptive,
    });
  });

  return {
    items,
    headerRow,
    columns: cols,
  };
}
