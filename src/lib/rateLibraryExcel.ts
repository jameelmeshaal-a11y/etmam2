import ExcelJS from 'exceljs';
import type { RateLibraryItem } from '../types';
import { BOQ_CATEGORIES, CATEGORY_LABELS_AR } from '../types';

const CATEGORY_AR_TO_KEY: Record<string, string> = Object.fromEntries(
  Object.entries(CATEGORY_LABELS_AR).map(([k, v]) => [v, k])
);

const CATEGORY_ALIAS_MAP: Record<string, string> = {
  'general': 'general',
  'عام': 'general',
  'عامة': 'general',
  'concrete': 'concrete',
  'خرسانة': 'concrete',
  'خرساني': 'concrete',
  'slab_concrete': 'concrete',
  'general_concrete': 'concrete',
  'foundation_concrete': 'concrete',
  'beam_concrete': 'concrete',
  'column_concrete': 'concrete',
  'blinding_concrete': 'concrete',
  'shear_wall_concrete': 'concrete',
  'Concrete': 'concrete',
  'blockwork': 'blockwork',
  'بلوك': 'blockwork',
  'بناء': 'blockwork',
  'finishes': 'finishes',
  'تشطيبات': 'finishes',
  'تشطيب': 'finishes',
  'Finishing': 'finishes',
  'tiling': 'finishes',
  'plastering': 'finishes',
  'cladding': 'finishes',
  'painting': 'painting',
  'دهانات': 'painting',
  'طلاء': 'painting',
  'excavation': 'excavation',
  'حفر': 'excavation',
  'حفر وردم': 'excavation',
  'أعمال ترابية': 'excavation',
  'Earthworks': 'excavation',
  'earthworks': 'excavation',
  'backfill': 'excavation',
  'asphalt': 'excavation',
  'steel': 'steel',
  'حديد': 'steel',
  'حديد تسليح': 'steel',
  'steel_misc': 'steel',
  'waterproofing': 'waterproofing',
  'عزل مائي': 'waterproofing',
  'Waterproofing': 'waterproofing',
  'insulation': 'insulation',
  'عزل حراري': 'insulation',
  'thermal_insulation': 'insulation',
  'plumbing': 'plumbing',
  'سباكة': 'plumbing',
  'صحية': 'plumbing',
  'Plumbing': 'plumbing',
  'plumbing_pipes': 'plumbing',
  'plumbing_fixtures': 'plumbing',
  'electrical': 'electrical',
  'كهرباء': 'electrical',
  'كهربائية': 'electrical',
  'Electrical': 'electrical',
  'electrical_wiring': 'electrical',
  'electrical_fixtures': 'electrical',
  'electrical_panels': 'electrical',
  'hvac': 'hvac',
  'تكييف': 'hvac',
  'تهوية': 'hvac',
  'ميكانيكية': 'hvac',
  'Mechanical': 'hvac',
  'hvac_equipment': 'hvac',
  'hvac_ductwork': 'hvac',
  'doors': 'doors_windows',
  'أبواب': 'doors_windows',
  'doors_windows': 'doors_windows',
  'Doors & Windows': 'doors_windows',
  'windows': 'doors_windows',
  'flooring': 'flooring',
  'أرضيات': 'flooring',
  'roofing': 'roofing',
  'أسقف': 'roofing',
  'fire_fighting': 'general',
  'Firefighting': 'general',
  'furniture': 'general',
  'slab concrete': 'concrete',
};

function resolveCategory(raw: string | undefined): string {
  if (!raw) return 'general';
  const trimmed = String(raw).trim();
  const direct = CATEGORY_ALIAS_MAP[trimmed];
  if (direct) return direct;
  if ((BOQ_CATEGORIES as readonly string[]).includes(trimmed)) return trimmed;
  const fromAr = CATEGORY_AR_TO_KEY[trimmed];
  if (fromAr) return fromAr;
  const lower = trimmed.toLowerCase();
  const lowerMatch = CATEGORY_ALIAS_MAP[lower];
  if (lowerMatch) return lowerMatch;
  const found = BOQ_CATEGORIES.find(c => c === lower);
  return found ?? 'general';
}

function cellStr(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && 'richText' in v) {
    return (v as ExcelJS.CellRichTextValue).richText.map(r => r.text).join('');
  }
  return String(v).trim();
}

function cellNum(cell: ExcelJS.Cell): number {
  const v = cell.value;
  if (v === null || v === undefined) return 0;
  const n = typeof v === 'number' ? v : parseFloat(String(v));
  return isNaN(n) ? 0 : n;
}

export interface ImportResult {
  success: number;
  failed: number;
  errors: string[];
  items: ParsedRateItem[];
}

export interface ParsedRateItem {
  standard_name_ar: string;
  standard_name_en: string;
  category: string;
  unit: string;
  rate_base: number;
  rate_target: number;
  rate_min: number;
  rate_max: number;
  source_type: 'Approved' | 'Field-Approved' | 'Draft';
  item_name_aliases?: string[];
}

function detectPriceLibraryFormat(headers: Record<number, string>): 'standard' | 'simplified' {
  const vals = Object.values(headers).join(' ');
  if (vals.includes('اسم البند') || vals.includes('الأسماء البديلة') || vals.includes('السعر')) {
    return 'simplified';
  }
  return 'standard';
}

export async function parseRateLibraryExcel(file: File): Promise<ParsedRateItem[]> {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('لا توجد أوراق عمل في الملف');

  const items: ParsedRateItem[] = [];

  const headerRow = sheet.getRow(1);
  const headers: Record<number, string> = {};
  headerRow.eachCell((cell, colNumber) => {
    headers[colNumber] = cellStr(cell).toLowerCase().trim();
  });

  const colIndex = (names: string[]): number => {
    for (const [col, header] of Object.entries(headers)) {
      for (const name of names) {
        if (header.includes(name)) return Number(col);
      }
    }
    return -1;
  };

  const format = detectPriceLibraryFormat(headers);

  if (format === 'simplified') {
    const colNameAr = colIndex(['اسم البند', 'اسم عربي', 'arabic', 'الاسم العربي']);
    const colNameEn = colIndex(['الاسم الإنجليزي', 'الاسم الانجليزي', 'english', 'en']);
    const colAliases = colIndex(['الأسماء البديلة', 'بديلة', 'alternative']);
    const colCategory = colIndex(['التصنيف', 'فئة', 'category', 'cat', 'الفئة']);
    const colUnit = colIndex(['الوحدة', 'وحدة', 'unit']);
    const colPrice = colIndex(['السعر', 'price', 'سعر']);
    const colSource = colIndex(['معتمد', 'approved', 'source', 'نوع']);

    if (colNameAr === -1 && colPrice === -1) {
      throw new Error('تعذّر العثور على أعمدة مكتبة الأسعار. تأكد أن الصف الأول يحتوي على رؤوس الأعمدة.');
    }

    const nameCol = colNameAr !== -1 ? colNameAr : 2;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const nameAr = cellStr(row.getCell(nameCol));
      if (!nameAr) return;

      const unit = colUnit !== -1 ? cellStr(row.getCell(colUnit)) || 'عدد' : 'عدد';
      const category = resolveCategory(colCategory !== -1 ? cellStr(row.getCell(colCategory)) : '');
      const nameEn = colNameEn !== -1 ? cellStr(row.getCell(colNameEn)) : '';
      const price = colPrice !== -1 ? cellNum(row.getCell(colPrice)) : 0;
      const aliasesRaw = colAliases !== -1 ? cellStr(row.getCell(colAliases)) : '';
      const aliases = aliasesRaw ? aliasesRaw.split(/[،,\n]+/).map(s => s.trim()).filter(Boolean) : [];

      let sourceType: 'Approved' | 'Field-Approved' | 'Draft' = 'Approved';
      if (colSource !== -1) {
        const raw = cellStr(row.getCell(colSource)).toLowerCase();
        if (raw === 'لا' || raw === 'no') sourceType = 'Draft';
        else if (raw.includes('field') || raw.includes('ميداني')) sourceType = 'Field-Approved';
        else if (raw.includes('draft') || raw.includes('مسودة')) sourceType = 'Draft';
      }

      items.push({
        standard_name_ar: nameAr,
        standard_name_en: nameEn,
        category,
        unit,
        rate_base: price,
        rate_target: price,
        rate_min: 0,
        rate_max: 0,
        source_type: sourceType,
        item_name_aliases: aliases,
      });
    });
  } else {
    const colNameAr = colIndex(['اسم عربي', 'arabic', 'ar', 'الاسم العربي', 'اسم_عربي', 'standard_name_ar', 'name_ar']);
    const colNameEn = colIndex(['اسم انجليزي', 'english', 'en', 'الاسم الانجليزي', 'standard_name_en', 'name_en']);
    const colCategory = colIndex(['فئة', 'category', 'cat', 'الفئة', 'التصنيف']);
    const colUnit = colIndex(['وحدة', 'unit', 'الوحدة']);
    const colBase = colIndex(['base', 'أساسي', 'rate_base', 'سعر_أساسي', 'سعر اساسي', 'base rate']);
    const colTarget = colIndex(['target', 'مستهدف', 'rate_target', 'سعر_مستهدف', 'سعر مستهدف', 'target rate']);
    const colMin = colIndex(['min', 'أدنى', 'rate_min', 'سعر_أدنى', 'سعر ادنى', 'min rate']);
    const colMax = colIndex(['max', 'أقصى', 'rate_max', 'سعر_أقصى', 'سعر اقصى', 'max rate']);
    const colSource = colIndex(['source', 'نوع', 'source_type', 'نوع_المصدر', 'مصدر']);

    if (colNameAr === -1) {
      throw new Error('تعذّر العثور على عمود الاسم العربي. تأكد أن الصف الأول يحتوي على رؤوس الأعمدة.');
    }

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const nameAr = cellStr(row.getCell(colNameAr));
      if (!nameAr) return;

      const unit = colUnit !== -1 ? cellStr(row.getCell(colUnit)) || 'm²' : 'm²';
      const category = resolveCategory(colCategory !== -1 ? cellStr(row.getCell(colCategory)) : '');
      const nameEn = colNameEn !== -1 ? cellStr(row.getCell(colNameEn)) : '';
      const rateBase = colBase !== -1 ? cellNum(row.getCell(colBase)) : 0;
      const rateTarget = colTarget !== -1 ? cellNum(row.getCell(colTarget)) : rateBase;
      const rateMin = colMin !== -1 ? cellNum(row.getCell(colMin)) : 0;
      const rateMax = colMax !== -1 ? cellNum(row.getCell(colMax)) : 0;

      let sourceType: 'Approved' | 'Field-Approved' | 'Draft' = 'Approved';
      if (colSource !== -1) {
        const raw = cellStr(row.getCell(colSource)).toLowerCase();
        if (raw.includes('field') || raw.includes('ميداني')) sourceType = 'Field-Approved';
        else if (raw.includes('draft') || raw.includes('مسودة')) sourceType = 'Draft';
      }

      items.push({ standard_name_ar: nameAr, standard_name_en: nameEn, category, unit, rate_base: rateBase, rate_target: rateTarget, rate_min: rateMin, rate_max: rateMax, source_type: sourceType });
    });
  }

  return items;
}

export async function exportRateLibraryToExcel(items: RateLibraryItem[]): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('مكتبة الأسعار');

  sheet.views = [{ rightToLeft: true }];

  const headers = [
    { header: 'كود البند', key: 'item_code', width: 14 },
    { header: 'اسم البند', key: 'standard_name_ar', width: 45 },
    { header: 'الاسم الإنجليزي', key: 'standard_name_en', width: 40 },
    { header: 'الأسماء البديلة', key: 'aliases', width: 50 },
    { header: 'التصنيف', key: 'category', width: 20 },
    { header: 'الوحدة', key: 'unit', width: 12 },
    { header: 'السعر', key: 'rate_base', width: 16 },
    { header: 'العملة', key: 'currency', width: 10 },
    { header: 'معتمد', key: 'approved', width: 10 },
  ];

  sheet.columns = headers;

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 24;

  items.forEach(item => {
    const row = sheet.addRow({
      item_code: '',
      standard_name_ar: item.standard_name_ar,
      standard_name_en: item.standard_name_en || '',
      aliases: (item.item_name_aliases ?? []).join('، '),
      category: CATEGORY_LABELS_AR[item.category] ?? item.category,
      unit: item.unit,
      rate_base: item.rate_base,
      currency: 'SAR',
      approved: item.source_type === 'Approved' || item.source_type === 'Field-Approved' ? 'نعم' : 'لا',
    });

    const priceCell = row.getCell('rate_base');
    priceCell.numFmt = '#,##0.00';
  });

  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: headers.length },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `مكتبة_الأسعار_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}
