import { X, Target, Tag, Ruler, TrendingUp, CheckCircle, AlertCircle } from 'lucide-react';
import type { BOQItem, RateLibraryItem } from '../types';
import { CATEGORY_LABELS_AR } from '../types';

interface MatchDetailsModalProps {
  item: BOQItem;
  libraryItem: RateLibraryItem | null;
  onClose: () => void;
}

const sarFormat = (v: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR' }).format(v);

function ConfidenceBar({ value }: { value: number }) {
  let colorClass = 'bg-red-500';
  let textClass = 'text-red-700';
  if (value >= 95) { colorClass = 'bg-emerald-500'; textClass = 'text-emerald-700'; }
  else if (value >= 85) { colorClass = 'bg-blue-500'; textClass = 'text-blue-700'; }
  else if (value >= 75) { colorClass = 'bg-amber-500'; textClass = 'text-amber-700'; }

  return (
    <div className="space-y-1.5">
      <div className="flex items-center justify-between">
        <span className="text-xs text-slate-500">درجة الثقة</span>
        <span className={`text-sm font-bold ${textClass}`}>{value}%</span>
      </div>
      <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
        <div
          className={`h-full rounded-full transition-all duration-500 ${colorClass}`}
          style={{ width: `${value}%` }}
        />
      </div>
    </div>
  );
}

const STAGE_LABELS: Record<number, { label: string; description: string }> = {
  1: { label: 'المرحلة 1', description: 'تطابق رقم البند (دقيق)' },
  2: { label: 'المرحلة 2', description: 'تصفية الوحدة والفئة' },
  3: { label: 'المرحلة 3', description: 'تطابق الوصف' },
  4: { label: 'المرحلة 4', description: 'التطابق المركّب' },
};

const SOURCE_TYPE_COLORS: Record<string, string> = {
  'Approved': 'bg-emerald-100 text-emerald-800 border-emerald-200',
  'Field-Approved': 'bg-blue-100 text-blue-800 border-blue-200',
  'Draft': 'bg-amber-100 text-amber-800 border-amber-200',
};

const SOURCE_TYPE_LABELS: Record<string, string> = {
  'Approved': 'معتمد',
  'Field-Approved': 'معتمد ميداني',
  'Draft': 'مسودة',
};

export default function MatchDetailsModal({ item, libraryItem, onClose }: MatchDetailsModalProps) {
  const stageInfo = STAGE_LABELS[item.confidence > 0 ? 3 : 0];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm" dir="rtl">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg fade-in">
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100">
          <div>
            <h2 className="text-base font-semibold text-slate-900">تفاصيل التطابق</h2>
            <p className="text-xs text-slate-500 mt-0.5">نتائج محرك المطابقة</p>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600"
          >
            <X size={18} />
          </button>
        </div>

        <div className="px-6 py-4 bg-slate-50 border-b border-slate-100">
          <p className="text-xs text-slate-500 mb-1">بند الـ BOQ</p>
          <p className="text-sm font-medium text-slate-800 leading-relaxed">{item.description}</p>
          <div className="flex items-center gap-3 mt-2">
            <span className="text-xs text-slate-500">الوحدة:</span>
            <span className="text-xs font-medium bg-slate-200 text-slate-700 px-2 py-0.5 rounded">{item.unit || '—'}</span>
            <span className="text-xs text-slate-500">رقم البند:</span>
            <span className="text-xs font-medium text-slate-700 ltr">{item.item_no || '—'}</span>
          </div>
        </div>

        {!libraryItem ? (
          <div className="px-6 py-8 text-center">
            <AlertCircle size={40} className="mx-auto text-slate-300 mb-3" />
            <p className="text-sm text-slate-500">لم يتم العثور على تطابق في مكتبة الأسعار</p>
            <p className="text-xs text-slate-400 mt-1">يتطلب إدخالاً يدوياً أو إضافة للمكتبة</p>
          </div>
        ) : (
          <div className="px-6 py-4 space-y-4">
            <div className="flex items-start gap-3">
              <CheckCircle size={18} className="text-emerald-500 mt-0.5 flex-shrink-0" />
              <div className="flex-1">
                <p className="text-sm font-semibold text-slate-900">{libraryItem.standard_name_ar}</p>
                {libraryItem.standard_name_en && (
                  <p className="text-xs text-slate-500 mt-0.5 ltr">{libraryItem.standard_name_en}</p>
                )}
              </div>
              <span className={`px-2 py-0.5 text-xs font-medium rounded-full border ${SOURCE_TYPE_COLORS[libraryItem.source_type] ?? 'bg-slate-100 text-slate-600'}`}>
                {SOURCE_TYPE_LABELS[libraryItem.source_type] ?? libraryItem.source_type}
              </span>
            </div>

            <ConfidenceBar value={item.confidence} />

            {stageInfo && (
              <div className="flex items-center gap-2 p-3 bg-blue-50 rounded-lg border border-blue-100">
                <Target size={16} className="text-blue-600 flex-shrink-0" />
                <div>
                  <span className="text-xs font-semibold text-blue-800">{stageInfo.label}</span>
                  <span className="text-xs text-blue-600 mr-2">— {stageInfo.description}</span>
                </div>
              </div>
            )}

            <div className="grid grid-cols-2 gap-3">
              <div className="flex items-center gap-2 p-3 bg-slate-50 rounded-lg">
                <Tag size={14} className="text-slate-400 flex-shrink-0" />
                <div>
                  <p className="text-xs text-slate-500">الفئة</p>
                  <p className="text-sm font-medium text-slate-800">
                    {CATEGORY_LABELS_AR[libraryItem.category] ?? libraryItem.category}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2 p-3 bg-slate-50 rounded-lg">
                <Ruler size={14} className="text-slate-400 flex-shrink-0" />
                <div>
                  <p className="text-xs text-slate-500">الوحدة</p>
                  <p className="text-sm font-medium text-slate-800">{libraryItem.unit}</p>
                </div>
              </div>
            </div>

            <div className="border border-slate-100 rounded-lg overflow-hidden">
              <div className="bg-slate-50 px-3 py-2 border-b border-slate-100">
                <div className="flex items-center gap-2">
                  <TrendingUp size={14} className="text-slate-500" />
                  <span className="text-xs font-semibold text-slate-700">نطاق الأسعار</span>
                </div>
              </div>
              <div className="grid grid-cols-4 divide-x divide-x-reverse divide-slate-100">
                {[
                  { label: 'الأساسي', value: libraryItem.rate_base, highlight: false },
                  { label: 'المستهدف', value: libraryItem.rate_target, highlight: true },
                  { label: 'الأدنى', value: libraryItem.rate_min, highlight: false },
                  { label: 'الأقصى', value: libraryItem.rate_max, highlight: false },
                ].map(({ label, value, highlight }) => (
                  <div key={label} className={`px-2 py-3 text-center ${highlight ? 'bg-emerald-50' : ''}`}>
                    <p className="text-xs text-slate-500 mb-1">{label}</p>
                    <p className={`text-xs font-bold ltr ${highlight ? 'text-emerald-700' : 'text-slate-700'}`}>
                      {sarFormat(value)}
                    </p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        <div className="px-6 py-4 border-t border-slate-100">
          <button
            onClick={onClose}
            className="w-full px-4 py-2.5 text-sm text-slate-600 hover:bg-slate-100 rounded-lg transition-colors font-medium"
          >
            إغلاق
          </button>
        </div>
      </div>
    </div>
  );
}
