export interface Project {
  id: string;
  name: string;
  client: string;
  city: string;
  created_by: string;
  created_at: string;
  updated_at: string;
  boq_files?: BOQFile[];
}

export interface BOQFile {
  id: string;
  project_id: string;
  name: string;
  storage_path: string;
  city: string;
  created_at: string;
  total_items: number;
  priced_items: number;
  export_variance_pct: number;
  created_by: string;
  is_archived: boolean;
  archived_at: string | null;
  notes: string;
}

export interface RateLibraryItem {
  id: string;
  standard_name_ar: string;
  standard_name_en: string;
  category: string;
  unit: string;
  rate_base: number;
  rate_target: number;
  rate_min: number;
  rate_max: number;
  keywords: string[];
  item_name_aliases: string[];
  is_locked: boolean;
  source_type: 'Approved' | 'Field-Approved' | 'Draft';
  approved_at: string | null;
  last_reviewed_at: string;
  created_by: string;
  created_at: string;
  updated_at: string;
}

export type BOQItemStatus = 'pending' | 'approved' | 'stale_price' | 'descriptive' | 'needs_review' | 'manual';

export interface BOQItem {
  id: string;
  boq_file_id: string;
  item_no: string;
  description: string;
  unit: string;
  quantity: number;
  unit_rate: number | null;
  total_price: number | null;
  status: BOQItemStatus;
  override_type: 'manual' | null;
  linked_rate_id: string | null;
  confidence: number;
  row_index: number;
  materials: number;
  labor: number;
  equipment: number;
  logistics: number;
  risk: number;
  profit: number;
  created_at: string;
  updated_at: string;
  category: string | null;
  matched_library_item?: RateLibraryItem;
}

export interface RateSource {
  id: string;
  boq_item_id: string;
  rate_library_id: string | null;
  unit_rate: number;
  source_type: string;
  override_type: string | null;
  materials: number;
  labor: number;
  equipment: number;
  logistics: number;
  risk: number;
  profit: number;
  created_by: string;
  created_at: string;
}

export interface MatchResult {
  libraryItem: RateLibraryItem;
  confidence: number;
  stage: 1 | 2 | 3 | 4;
  stageLabel: string;
}

export interface PriceBreakdown {
  materials: number;
  labor: number;
  equipment: number;
  logistics: number;
  risk: number;
  profit: number;
}

export interface ExportResult {
  success: boolean;
  injected: number;
  total: number;
  variance: number;
  unmatched: string[];
  error?: string;
}

export const BOQ_CATEGORIES = [
  'concrete',
  'blockwork',
  'finishes',
  'excavation',
  'steel',
  'waterproofing',
  'insulation',
  'plumbing',
  'electrical',
  'hvac',
  'doors_windows',
  'flooring',
  'painting',
  'roofing',
  'general',
] as const;

export type BOQCategory = typeof BOQ_CATEGORIES[number];

export const CATEGORY_LABELS_AR: Record<string, string> = {
  concrete: 'خرسانة',
  blockwork: 'بلوك',
  finishes: 'تشطيبات',
  excavation: 'حفر وردم',
  steel: 'حديد تسليح',
  waterproofing: 'عزل مائي',
  insulation: 'عزل حراري',
  plumbing: 'سباكة',
  electrical: 'كهرباء',
  hvac: 'تكييف وتهوية',
  doors_windows: 'أبواب ونوافذ',
  flooring: 'أرضيات',
  painting: 'دهانات',
  roofing: 'أسقف',
  general: 'عام',
};

export const STATUS_LABELS_AR: Record<BOQItemStatus, string> = {
  pending: 'في الانتظار',
  approved: 'معتمد',
  stale_price: 'سعر قديم',
  descriptive: 'وصفي',
  needs_review: 'يحتاج مراجعة',
  manual: 'يدوي',
};

export const UNIT_LABELS = [
  'm²', 'm³', 'm', 'lm', 'nos', 'kg', 'ton', 'ls', 'set', 'item',
  'متر مربع', 'متر مكعب', 'متر طولي', 'عدد', 'طن', 'كجم', 'مجموعة',
];
