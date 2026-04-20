import { useState, useEffect, useCallback, useRef } from 'react';
import { Plus, CreditCard as Edit2, Trash2, Lock, Unlock, Search, X, Save, Upload, Download, AlertCircle, CheckCircle2 } from 'lucide-react';
import { getAllLibraryItems, upsertLibraryItem, deleteLibraryItem, lockLibraryItem, bulkImportLibraryItems, lockAllLibraryItems } from '../lib/priceLibrary';
import { parseRateLibraryExcel, exportRateLibraryToExcel } from '../lib/rateLibraryExcel';
import type { RateLibraryItem } from '../types';
import { BOQ_CATEGORIES, CATEGORY_LABELS_AR, UNIT_LABELS } from '../types';

interface RateLibraryTableProps {
  onClose?: () => void;
}

const SOURCE_TYPE_OPTIONS = ['Approved', 'Field-Approved', 'Draft'] as const;
type SourceType = typeof SOURCE_TYPE_OPTIONS[number];

const SOURCE_LABELS: Record<SourceType, string> = {
  'Approved': 'معتمد',
  'Field-Approved': 'معتمد ميداني',
  'Draft': 'مسودة',
};

const SOURCE_COLORS: Record<SourceType, string> = {
  'Approved': 'bg-emerald-100 text-emerald-800 border-emerald-200',
  'Field-Approved': 'bg-blue-100 text-blue-800 border-blue-200',
  'Draft': 'bg-amber-100 text-amber-800 border-amber-200',
};

const sarFormat = (v: number) =>
  new Intl.NumberFormat('ar-SA', { style: 'currency', currency: 'SAR', maximumFractionDigits: 0 }).format(v);

interface FormData {
  id?: string;
  standard_name_ar: string;
  standard_name_en: string;
  category: string;
  unit: string;
  rate_base: string;
  rate_target: string;
  rate_min: string;
  rate_max: string;
  source_type: SourceType;
}

const emptyForm = (): FormData => ({
  standard_name_ar: '',
  standard_name_en: '',
  category: 'general',
  unit: 'm²',
  rate_base: '',
  rate_target: '',
  rate_min: '',
  rate_max: '',
  source_type: 'Draft',
});

function RateFormModal({
  initial,
  onSave,
  onClose,
}: {
  initial: FormData;
  onSave: (data: FormData) => Promise<void>;
  onClose: () => void;
}) {
  const [form, setForm] = useState<FormData>(initial);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  function set(key: keyof FormData, value: string) {
    setForm(prev => ({ ...prev, [key]: value }));
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!form.standard_name_ar.trim()) { setError('الاسم العربي مطلوب'); return; }
    if (!form.unit.trim()) { setError('الوحدة مطلوبة'); return; }
    setSaving(true);
    setError(null);
    try {
      await onSave(form);
      onClose();
    } catch (ex) {
      setError((ex as Error).message);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm" dir="rtl">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg fade-in max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 sticky top-0 bg-white z-10">
          <h2 className="text-base font-semibold text-slate-900">{form.id ? 'تعديل معدل' : 'إضافة معدل جديد'}</h2>
          <button onClick={onClose} className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100">
            <X size={18} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="px-6 py-4 space-y-4">
          <div className="grid grid-cols-1 gap-4">
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">الاسم العربي *</label>
              <input value={form.standard_name_ar} onChange={e => set('standard_name_ar', e.target.value)} className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="مثال: خرسانة عادية" />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">الاسم الإنجليزي</label>
              <input value={form.standard_name_en} onChange={e => set('standard_name_en', e.target.value)} className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 ltr" placeholder="e.g. Plain Concrete" dir="ltr" />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">الفئة</label>
              <select value={form.category} onChange={e => set('category', e.target.value)} className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                {BOQ_CATEGORIES.map(cat => (
                  <option key={cat} value={cat}>{CATEGORY_LABELS_AR[cat]}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">الوحدة *</label>
              <select value={form.unit} onChange={e => set('unit', e.target.value)} className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                {UNIT_LABELS.map(u => <option key={u} value={u}>{u}</option>)}
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 gap-4">
            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">السعر (ر.س) *</label>
              <input
                type="number"
                min="0"
                step="0.01"
                value={form.rate_base}
                onChange={e => { set('rate_base', e.target.value); set('rate_target', e.target.value); }}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 ltr text-right"
                placeholder="0.00"
              />
            </div>
          </div>
          <details className="group">
            <summary className="text-xs text-slate-500 cursor-pointer hover:text-slate-700 flex items-center gap-1 select-none">
              <span className="group-open:rotate-90 transition-transform inline-block">›</span>
              خيارات متقدمة (الحد الأدنى / الأقصى)
            </summary>
            <div className="grid grid-cols-2 gap-4 mt-3">
              {[
                { key: 'rate_min' as keyof FormData, label: 'الحد الأدنى' },
                { key: 'rate_max' as keyof FormData, label: 'الحد الأقصى' },
              ].map(({ key, label }) => (
                <div key={key}>
                  <label className="block text-xs font-medium text-slate-700 mb-1">{label} (ر.س)</label>
                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={form[key] as string}
                    onChange={e => set(key, e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 ltr text-right"
                    placeholder="0.00"
                  />
                </div>
              ))}
            </div>
          </details>

          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">نوع المصدر</label>
            <div className="flex items-center gap-2">
              {SOURCE_TYPE_OPTIONS.map(st => (
                <button
                  key={st}
                  type="button"
                  onClick={() => set('source_type', st)}
                  className={`flex-1 py-2 text-xs font-medium rounded-lg border transition-colors ${
                    form.source_type === st ? SOURCE_COLORS[st] : 'bg-slate-50 text-slate-500 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {SOURCE_LABELS[st]}
                </button>
              ))}
            </div>
          </div>

          {error && (
            <div className="p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">{error}</div>
          )}

          <div className="flex items-center gap-3 pt-2">
            <button type="submit" disabled={saving} className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-medium rounded-lg transition-colors">
              <Save size={16} />
              <span>{saving ? 'جاري الحفظ...' : 'حفظ'}</span>
            </button>
            <button type="button" onClick={onClose} className="px-4 py-2.5 text-sm text-slate-600 hover:bg-slate-100 rounded-lg transition-colors">إلغاء</button>
          </div>
        </form>
      </div>
    </div>
  );
}

interface ImportStatus {
  type: 'success' | 'error' | 'partial';
  message: string;
}

const ROW_HEIGHT = 52;
const OVERSCAN = 10;

function VirtualTable({
  items,
  onEdit,
  onDelete,
  onLock,
  deletingId,
}: {
  items: RateLibraryItem[];
  onEdit: (item: RateLibraryItem) => void;
  onDelete: (id: string) => void;
  onLock: (id: string, lock: boolean) => void;
  deletingId: string | null;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [containerHeight, setContainerHeight] = useState(600);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(entries => {
      setContainerHeight(entries[0].contentRect.height);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const totalHeight = items.length * ROW_HEIGHT;
  const startIndex = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN);
  const visibleCount = Math.ceil(containerHeight / ROW_HEIGHT) + OVERSCAN * 2;
  const endIndex = Math.min(items.length - 1, startIndex + visibleCount);

  const visibleItems = items.slice(startIndex, endIndex + 1);

  return (
    <div
      ref={containerRef}
      className="flex-1 overflow-auto"
      onScroll={e => setScrollTop((e.target as HTMLDivElement).scrollTop)}
      style={{ position: 'relative' }}
    >
      <table className="w-full text-sm min-w-[1120px] border-collapse">
        <thead className="bg-slate-50 sticky top-0 z-10">
          <tr>
            <th className="px-3 py-3 text-xs font-semibold text-slate-400 border-b border-slate-200 w-12 text-center">#</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 text-right">اسم البند</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 text-right">الأسماء البديلة</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 w-28 text-right">التصنيف</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 w-16 text-center">الوحدة</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 w-32 text-left">السعر (ر.س)</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 w-24 text-center">الحالة</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 w-16 text-center">مقفل</th>
            <th className="px-3 py-3 text-xs font-semibold text-slate-500 border-b border-slate-200 w-20 text-center">إجراءات</th>
          </tr>
        </thead>
        <tbody className="bg-white" style={{ position: 'relative' }}>
          <tr style={{ height: startIndex * ROW_HEIGHT }}>
            <td colSpan={9} />
          </tr>
          {visibleItems.map((item, relIdx) => {
            const absIndex = startIndex + relIdx;
            return (
              <tr
                key={item.id}
                className="hover:bg-slate-50/80 transition-colors group border-b border-slate-100"
                style={{ height: ROW_HEIGHT }}
              >
                <td className="px-3 py-2 text-xs text-slate-400 text-center font-mono select-none tabular-nums">
                  {absIndex + 1}
                </td>
                <td className="px-3 py-2 max-w-xs">
                  <p className="text-sm font-medium text-slate-800 leading-snug truncate">{item.standard_name_ar}</p>
                  {item.standard_name_en && (
                    <p className="text-xs text-slate-400 ltr mt-0.5 truncate">{item.standard_name_en}</p>
                  )}
                </td>
                <td className="px-3 py-2 max-w-xs">
                  <p className="text-xs text-slate-500 leading-relaxed truncate">
                    {item.item_name_aliases?.length ? item.item_name_aliases.slice(0, 3).join('، ') : '—'}
                  </p>
                </td>
                <td className="px-3 py-2">
                  <span className="inline-flex items-center px-2 py-0.5 text-xs font-medium rounded-md bg-slate-100 text-slate-600 truncate max-w-full">
                    {CATEGORY_LABELS_AR[item.category] ?? item.category}
                  </span>
                </td>
                <td className="px-3 py-2 text-xs text-center text-slate-600 font-medium">{item.unit}</td>
                <td className="px-3 py-2 text-sm ltr text-left font-semibold text-slate-800">
                  {sarFormat(item.rate_base)}
                </td>
                <td className="px-3 py-2 text-center">
                  <span className={`inline-flex px-2 py-0.5 text-xs font-medium rounded-full border ${SOURCE_COLORS[item.source_type] ?? 'bg-slate-100 text-slate-600'}`}>
                    {SOURCE_LABELS[item.source_type] ?? item.source_type}
                  </span>
                </td>
                <td className="px-3 py-2 text-center">
                  <button
                    onClick={() => onLock(item.id, !item.is_locked)}
                    title={item.is_locked ? 'فتح القفل' : 'قفل'}
                    className={`w-7 h-7 rounded-md flex items-center justify-center mx-auto transition-colors ${item.is_locked ? 'text-amber-600 hover:bg-amber-50' : 'text-slate-400 hover:bg-slate-100'}`}
                  >
                    {item.is_locked ? <Lock size={14} /> : <Unlock size={14} />}
                  </button>
                </td>
                <td className="px-3 py-2">
                  <div className="flex items-center justify-center gap-1">
                    {!item.is_locked && (
                      <>
                        <button onClick={() => onEdit(item)} title="تعديل" className="w-7 h-7 rounded-md flex items-center justify-center text-slate-400 hover:bg-blue-50 hover:text-blue-600 transition-colors">
                          <Edit2 size={13} />
                        </button>
                        <button onClick={() => onDelete(item.id)} disabled={deletingId === item.id} title="حذف" className="w-7 h-7 rounded-md flex items-center justify-center text-slate-400 hover:bg-red-50 hover:text-red-600 transition-colors disabled:opacity-40">
                          <Trash2 size={13} />
                        </button>
                      </>
                    )}
                  </div>
                </td>
              </tr>
            );
          })}
          <tr style={{ height: Math.max(0, (items.length - endIndex - 1) * ROW_HEIGHT) }}>
            <td colSpan={9} />
          </tr>
        </tbody>
      </table>

      {items.length === 0 && (
        <div className="flex flex-col items-center justify-center py-20 text-slate-400">
          <Search size={32} className="mb-3 opacity-40" />
          <p className="text-sm">لا توجد نتائج مطابقة</p>
        </div>
      )}

      <div
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          width: 1,
          height: totalHeight + 48,
          pointerEvents: 'none',
        }}
      />
    </div>
  );
}

export default function RateLibraryTable({ onClose }: RateLibraryTableProps) {
  const [items, setItems] = useState<RateLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [filterCategory, setFilterCategory] = useState('');
  const [filterSource, setFilterSource] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingItem, setEditingItem] = useState<RateLibraryItem | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);
  const [importing, setImporting] = useState(false);
  const [importProgress, setImportProgress] = useState<{ done: number; total: number } | null>(null);
  const [exporting, setExporting] = useState(false);
  const [importStatus, setImportStatus] = useState<ImportStatus | null>(null);
  const [lockingAll, setLockingAll] = useState(false);
  const [lockAllConfirm, setLockAllConfirm] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadItems = useCallback(async () => {
    setLoading(true);
    try {
      const data = await getAllLibraryItems();
      setItems(data);
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { loadItems(); }, [loadItems]);

  const filtered = items.filter(item => {
    if (search) {
      const q = search.toLowerCase();
      if (!item.standard_name_ar.includes(search) && !item.standard_name_en.toLowerCase().includes(q)) return false;
    }
    if (filterCategory && item.category !== filterCategory) return false;
    if (filterSource && item.source_type !== filterSource) return false;
    return true;
  });

  async function handleSave(form: FormData) {
    await upsertLibraryItem({
      id: form.id,
      standard_name_ar: form.standard_name_ar,
      standard_name_en: form.standard_name_en,
      category: form.category,
      unit: form.unit,
      rate_base: parseFloat(form.rate_base) || 0,
      rate_target: parseFloat(form.rate_target) || 0,
      rate_min: parseFloat(form.rate_min) || 0,
      rate_max: parseFloat(form.rate_max) || 0,
      source_type: form.source_type,
    });
    await loadItems();
  }

  async function handleDelete(id: string) {
    setDeletingId(id);
    setActionError(null);
    try {
      await deleteLibraryItem(id);
      await loadItems();
    } catch (e) {
      setActionError((e as Error).message);
    } finally {
      setDeletingId(null);
    }
  }

  async function handleLock(id: string, lock: boolean) {
    setActionError(null);
    try {
      await lockLibraryItem(id, lock);
      await loadItems();
    } catch (e) {
      setActionError((e as Error).message);
    }
  }

  async function handleLockAll(lock: boolean) {
    setLockingAll(true);
    setActionError(null);
    setLockAllConfirm(false);
    try {
      await lockAllLibraryItems(lock);
      await loadItems();
    } catch (e) {
      setActionError((e as Error).message);
    } finally {
      setLockingAll(false);
    }
  }

  async function handleImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';
    setImporting(true);
    setImportStatus(null);
    setImportProgress(null);
    setActionError(null);
    try {
      const parsed = await parseRateLibraryExcel(file);
      if (parsed.length === 0) {
        setImportStatus({ type: 'error', message: 'لم يتم العثور على بيانات في الملف.' });
        return;
      }
      const result = await bulkImportLibraryItems(parsed, (done, total) => {
        setImportProgress({ done, total });
      });
      await loadItems();
      setImportProgress(null);
      if (result.failed === 0) {
        setImportStatus({ type: 'success', message: `تم استيراد ${result.success} بند بنجاح.` });
      } else {
        setImportStatus({ type: 'partial', message: `تم استيراد ${result.success} بند. فشل: ${result.failed}. ${result.errors[0] ?? ''}` });
      }
    } catch (ex) {
      setImportProgress(null);
      setImportStatus({ type: 'error', message: (ex as Error).message });
    } finally {
      setImporting(false);
    }
  }

  async function handleExport() {
    setExporting(true);
    try {
      await exportRateLibraryToExcel(items);
    } catch (ex) {
      setActionError((ex as Error).message);
    } finally {
      setExporting(false);
    }
  }

  function openEdit(item: RateLibraryItem) {
    setEditingItem(item);
    setShowForm(true);
  }

  function formInitial(): FormData {
    if (editingItem) {
      return {
        id: editingItem.id,
        standard_name_ar: editingItem.standard_name_ar,
        standard_name_en: editingItem.standard_name_en,
        category: editingItem.category,
        unit: editingItem.unit,
        rate_base: String(editingItem.rate_base),
        rate_target: String(editingItem.rate_target),
        rate_min: String(editingItem.rate_min),
        rate_max: String(editingItem.rate_max),
        source_type: editingItem.source_type,
      };
    }
    return emptyForm();
  }

  const lockedCount = items.filter(i => i.is_locked).length;
  const allLocked = items.length > 0 && lockedCount === items.length;

  return (
    <div className="flex flex-col h-full" dir="rtl">
      <div className="bg-white border-b border-slate-200 px-6 py-4 flex-shrink-0">
        <div className="flex items-center justify-between mb-4">
          <div>
            <h2 className="text-lg font-semibold text-slate-900">مكتبة الأسعار</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              {filtered.length !== items.length
                ? <><span className="text-blue-600 font-medium">{filtered.length}</span> نتيجة من {items.length} بند · {lockedCount} مقفل</>
                : <>{items.length} بند · {lockedCount} مقفل</>
              }
            </p>
          </div>
          <div className="flex items-center gap-2 flex-wrap justify-end">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={handleImport}
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={importing}
              title="استيراد من Excel"
              className="flex items-center gap-2 px-3 py-2 border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
            >
              <Upload size={15} />
              <span>{importing ? 'جاري الاستيراد...' : 'استيراد Excel'}</span>
            </button>
            <button
              onClick={handleExport}
              disabled={exporting || items.length === 0}
              title="تصدير إلى Excel"
              className="flex items-center gap-2 px-3 py-2 border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-sm font-medium rounded-lg transition-colors disabled:opacity-50"
            >
              <Download size={15} />
              <span>{exporting ? 'جاري التصدير...' : 'تصدير Excel'}</span>
            </button>

            {lockAllConfirm ? (
              <div className="flex items-center gap-1 bg-amber-50 border border-amber-200 rounded-lg px-2 py-1">
                <span className="text-xs text-amber-800">
                  {allLocked ? 'فتح قفل الجميع؟' : 'قفل الجميع؟'}
                </span>
                <button
                  onClick={() => handleLockAll(!allLocked)}
                  disabled={lockingAll}
                  className="text-xs font-medium text-amber-900 hover:text-amber-700 px-1"
                >
                  {lockingAll ? '...' : 'نعم'}
                </button>
                <button onClick={() => setLockAllConfirm(false)} className="text-xs text-slate-500 hover:text-slate-700 px-1">لا</button>
              </div>
            ) : (
              <button
                onClick={() => setLockAllConfirm(true)}
                disabled={items.length === 0 || lockingAll}
                title={allLocked ? 'فتح قفل جميع الأسعار' : 'قفل جميع الأسعار'}
                className={`flex items-center gap-2 px-3 py-2 border text-sm font-medium rounded-lg transition-colors disabled:opacity-50 ${
                  allLocked
                    ? 'border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-700'
                    : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                }`}
              >
                {allLocked ? <Unlock size={15} /> : <Lock size={15} />}
                <span>{allLocked ? 'فتح قفل الجميع' : 'قفل الجميع'}</span>
              </button>
            )}

            <button
              onClick={() => { setEditingItem(null); setShowForm(true); }}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg transition-colors"
            >
              <Plus size={16} />
              <span>إضافة معدل</span>
            </button>
            {onClose && (
              <button onClick={onClose} className="w-9 h-9 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100">
                <X size={18} />
              </button>
            )}
          </div>
        </div>

        {actionError && (
          <div className="mb-3 p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">{actionError}</div>
        )}

        {importing && importProgress && (
          <div className="mb-3 p-3 bg-blue-50 border border-blue-200 rounded-lg">
            <div className="flex items-center justify-between text-xs text-blue-700 mb-1.5">
              <span>جاري استيراد البنود...</span>
              <span>{importProgress.done} / {importProgress.total}</span>
            </div>
            <div className="w-full bg-blue-100 rounded-full h-1.5">
              <div
                className="bg-blue-500 h-1.5 rounded-full transition-all duration-300"
                style={{ width: `${Math.round((importProgress.done / importProgress.total) * 100)}%` }}
              />
            </div>
          </div>
        )}

        {importStatus && (
          <div className={`mb-3 p-3 rounded-lg text-sm flex items-start gap-2 ${
            importStatus.type === 'success'
              ? 'bg-emerald-50 border border-emerald-200 text-emerald-800'
              : importStatus.type === 'partial'
              ? 'bg-amber-50 border border-amber-200 text-amber-800'
              : 'bg-red-50 border border-red-200 text-red-700'
          }`}>
            {importStatus.type === 'success'
              ? <CheckCircle2 size={16} className="flex-shrink-0 mt-0.5" />
              : <AlertCircle size={16} className="flex-shrink-0 mt-0.5" />}
            <span>{importStatus.message}</span>
            <button onClick={() => setImportStatus(null)} className="mr-auto flex-shrink-0 opacity-60 hover:opacity-100">
              <X size={14} />
            </button>
          </div>
        )}

        <div className="flex items-center gap-3 flex-wrap">
          <div className="flex items-center gap-2 flex-1 min-w-48 border border-slate-200 rounded-lg px-3 py-2 bg-white">
            <Search size={15} className="text-slate-400 flex-shrink-0" />
            <input
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="بحث بالاسم..."
              className="flex-1 text-sm focus:outline-none"
            />
            {search && (
              <button onClick={() => setSearch('')} className="text-slate-400 hover:text-slate-600">
                <X size={14} />
              </button>
            )}
          </div>
          <select value={filterCategory} onChange={e => setFilterCategory(e.target.value)} className="px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="">كل الفئات</option>
            {BOQ_CATEGORIES.map(cat => <option key={cat} value={cat}>{CATEGORY_LABELS_AR[cat]}</option>)}
          </select>
          <select value={filterSource} onChange={e => setFilterSource(e.target.value)} className="px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="">كل الأنواع</option>
            {SOURCE_TYPE_OPTIONS.map(st => <option key={st} value={st}>{SOURCE_LABELS[st]}</option>)}
          </select>
          {(search || filterCategory || filterSource) && (
            <button
              onClick={() => { setSearch(''); setFilterCategory(''); setFilterSource(''); }}
              className="flex items-center gap-1.5 px-3 py-2 text-sm text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-lg transition-colors"
            >
              <X size={14} />
              <span>مسح الفلاتر</span>
            </button>
          )}
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-20 flex-1">
          <div className="w-8 h-8 border-3 border-blue-500 border-t-transparent rounded-full animate-spin" />
        </div>
      ) : (
        <VirtualTable
          items={filtered}
          onEdit={openEdit}
          onDelete={handleDelete}
          onLock={handleLock}
          deletingId={deletingId}
        />
      )}

      <div className="bg-white border-t border-slate-100 px-6 py-2 flex-shrink-0 flex items-center justify-between">
        <span className="text-xs text-slate-400">
          {filtered.length > 0
            ? `عرض ${filtered.length} بند${filtered.length !== items.length ? ` (مفلتر من ${items.length})` : ''}`
            : 'لا توجد نتائج'}
        </span>
        <span className="text-xs text-slate-400">قائمة تمرير مستمرة · مرر لأعلى أو لأسفل للتنقل</span>
      </div>

      {showForm && (
        <RateFormModal
          initial={formInitial()}
          onSave={handleSave}
          onClose={() => { setShowForm(false); setEditingItem(null); }}
        />
      )}
    </div>
  );
}
