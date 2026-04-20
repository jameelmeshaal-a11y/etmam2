import { useState, useEffect, useCallback, useRef } from 'react';
import {
  RefreshCw, Eye, CreditCard as Edit2, Lock, ChevronRight,
  Download, AlertTriangle, CheckSquare, CheckCircle, XCircle,
  HelpCircle, FileDown, ThumbsUp, Check,
} from 'lucide-react';
import { supabase } from '../lib/supabase';
import { getAllLibraryItems } from '../lib/priceLibrary';
import { priceItemsSequentially, repriceSingleItem } from '../lib/pricingEngine';
import { exportBOQ, exportUnpricedItemsForLibrary } from '../lib/approvalExporter';
import type { BOQItem, BOQFile, RateLibraryItem, BOQItemStatus } from '../types';
import { CATEGORY_LABELS_AR } from '../types';
import StatusBadge from './StatusBadge';
import PriceBreakdownModal from './PriceBreakdownModal';
import MatchDetailsModal from './MatchDetailsModal';
import type { PricingProgress } from '../lib/pricingEngine';

interface BOQTableProps {
  boqFileId: string;
  boqFile: BOQFile;
  onBack: () => void;
}

const sarFormat = (v: number | null) =>
  v == null ? '—' : new Intl.NumberFormat('ar-SA', { style: 'currency', currency: 'SAR' }).format(v);

function confidenceColor(c: number): string {
  if (c >= 95) return 'text-emerald-700 bg-emerald-50';
  if (c >= 85) return 'text-blue-700 bg-blue-50';
  if (c >= 75) return 'text-amber-700 bg-amber-50';
  return 'text-red-700 bg-red-50';
}

type FilterStatus = 'all' | BOQItemStatus | 'unpriced';
type SortMode = 'default' | 'unit_rate_desc' | 'total_price_desc';

const FILTER_OPTIONS: { value: FilterStatus; label: string }[] = [
  { value: 'all', label: 'الكل' },
  { value: 'pending', label: 'في الانتظار' },
  { value: 'unpriced', label: 'غير مسعَّر' },
  { value: 'approved', label: 'معتمد' },
  { value: 'manual', label: 'يدوي' },
  { value: 'needs_review', label: 'يحتاج مراجعة' },
  { value: 'stale_price', label: 'سعر قديم' },
];

interface PricingProgressState {
  active: boolean;
  progress: PricingProgress | null;
  done: boolean;
  result: { priced: number; failed: number; total: number } | null;
  error: string | null;
}

export default function BOQTable({ boqFileId, boqFile, onBack }: BOQTableProps) {
  const [items, setItems] = useState<BOQItem[]>([]);
  const [library, setLibrary] = useState<RateLibraryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [filterStatus, setFilterStatus] = useState<FilterStatus>('all');
  const [search, setSearch] = useState('');
  const [hideDescriptive, setHideDescriptive] = useState(false);
  const [sortMode, setSortMode] = useState<SortMode>('default');
  const [pricingState, setPricingState] = useState<PricingProgressState>({
    active: false, progress: null, done: false, result: null, error: null,
  });
  const [exporting, setExporting] = useState(false);
  const [exportingUnpriced, setExportingUnpriced] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [repricingItemId, setRepricingItemId] = useState<string | null>(null);
  const [approvingItemId, setApprovingItemId] = useState<string | null>(null);
  const [editItem, setEditItem] = useState<BOQItem | null>(null);
  const [matchItemState, setMatchItemState] = useState<BOQItem | null>(null);
  const itemsRef = useRef<BOQItem[]>([]);

  // Silent refresh — no loading spinner, no white flash
  const refreshItems = useCallback(async () => {
    const { data, error } = await supabase
      .from('boq_items')
      .select('*')
      .eq('boq_file_id', boqFileId)
      .order('row_index', { ascending: true });
    if (!error && data) {
      itemsRef.current = data as BOQItem[];
      setItems(data as BOQItem[]);
    }
  }, [boqFileId]);

  // Initial load — shows skeleton
  const loadItems = useCallback(async () => {
    setLoading(true);
    setLoadError(null);
    try {
      const { data, error } = await supabase
        .from('boq_items')
        .select('*')
        .eq('boq_file_id', boqFileId)
        .order('row_index', { ascending: true });
      if (error) {
        setLoadError(`خطأ في تحميل البنود: ${error.message}`);
        return;
      }
      const rows = (data ?? []) as BOQItem[];
      itemsRef.current = rows;
      setItems(rows);
    } catch (e) {
      setLoadError(`خطأ: ${(e as Error).message}`);
    } finally {
      setLoading(false);
    }
  }, [boqFileId]);

  useEffect(() => {
    loadItems();
    getAllLibraryItems().then(setLibrary).catch(console.error);
  }, [loadItems]);

  // "Priceable" = not descriptive AND quantity > 0
  const priceableItems = items.filter(i => i.status !== 'descriptive' && (i.quantity ?? 0) > 0);
  // "Unpriced" = priceable with no unit_rate
  const unpricedItems = priceableItems.filter(i => i.unit_rate == null || i.unit_rate === 0);
  const unpricedCount = unpricedItems.length;
  const pricedCount = priceableItems.filter(i => i.unit_rate != null && i.unit_rate > 0).length;
  const pricedPct = priceableItems.length > 0 ? Math.round((pricedCount / priceableItems.length) * 100) : 0;
  const totalSAR = items.reduce((sum, i) => sum + (i.total_price ?? 0), 0);
  // pending = has price but confidence < 75, needs review/approval
  const pendingCount = items.filter(i => i.status === 'pending' || i.status === 'stale_price').length;
  // truly needs auto-pricing = pending AND no unit_rate AND qty > 0
  const needsPricingCount = items.filter(
    i => (i.status === 'pending' || i.status === 'stale_price') && (i.unit_rate == null) && (i.quantity ?? 0) > 0
  ).length;

  const filteredItems = (() => {
    let result = items.filter(item => {
      if (hideDescriptive && item.status === 'descriptive') return false;
      if (filterStatus === 'unpriced') {
        if (item.status === 'descriptive') return false;
        if ((item.quantity ?? 0) <= 0) return false;
        if (item.unit_rate != null && item.unit_rate > 0) return false;
        return true;
      }
      if (filterStatus !== 'all' && item.status !== filterStatus) return false;
      if (search && !item.description.includes(search) && !item.item_no.includes(search)) return false;
      return true;
    });

    if (sortMode === 'unit_rate_desc') {
      result = [...result].sort((a, b) => (b.unit_rate ?? 0) - (a.unit_rate ?? 0));
    } else if (sortMode === 'total_price_desc') {
      result = [...result].sort((a, b) => (b.total_price ?? 0) - (a.total_price ?? 0));
    }

    return result;
  })();

  async function handleBulkPrice() {
    if (pricingState.active) return;
    setPricingState({ active: true, progress: null, done: false, result: null, error: null });

    try {
      const result = await priceItemsSequentially(boqFileId, library, (prog) => {
        setPricingState(prev => ({ ...prev, active: true, progress: prog }));
      });

      // Silent refresh after pricing completes — no white flash
      await refreshItems();

      setPricingState({
        active: false,
        progress: null,
        done: true,
        result: { priced: result.priced, failed: result.failed, total: result.total },
        error: null,
      });
    } catch (e) {
      const msg = (e as Error).message ?? 'حدث خطأ أثناء التسعير';
      setPricingState({ active: false, progress: null, done: false, result: null, error: msg });
    }
  }

  async function handleRepriceSingle(itemId: string) {
    setRepricingItemId(itemId);
    try {
      await repriceSingleItem(itemId, library);
      // Update only this one item in state — no full reload
      const { data } = await supabase
        .from('boq_items')
        .select('*')
        .eq('id', itemId)
        .maybeSingle();
      if (data) {
        setItems(prev => prev.map(i => i.id === itemId ? data as BOQItem : i));
      }
    } catch (e) {
      console.error(e);
    } finally {
      setRepricingItemId(null);
    }
  }

  async function handleExport() {
    setExporting(true);
    setExportError(null);
    try {
      const result = await exportBOQ(boqFile, items);
      if (!result.success) setExportError(result.error ?? 'فشل التصدير');
    } catch (e) {
      setExportError((e as Error).message);
    } finally {
      setExporting(false);
    }
  }

  async function handleExportUnpriced() {
    setExportingUnpriced(true);
    try {
      await exportUnpricedItemsForLibrary(boqFile.name, items);
    } catch (e) {
      setExportError((e as Error).message);
    } finally {
      setExportingUnpriced(false);
    }
  }

  async function handleApprovePending() {
    const pendingWithPrice = items.filter(
      i => i.status === 'pending' && i.unit_rate != null && i.unit_rate > 0
    );
    if (pendingWithPrice.length === 0) return;
    const ids = pendingWithPrice.map(i => i.id);
    await supabase
      .from('boq_items')
      .update({ status: 'approved', updated_at: new Date().toISOString() })
      .in('id', ids);
    await refreshItems();
  }

  async function handleApproveSingle(itemId: string) {
    setApprovingItemId(itemId);
    try {
      await supabase
        .from('boq_items')
        .update({ status: 'approved', updated_at: new Date().toISOString() })
        .eq('id', itemId);
      setItems(prev => prev.map(i => i.id === itemId ? { ...i, status: 'approved' as BOQItemStatus } : i));
    } finally {
      setApprovingItemId(null);
    }
  }

  function getLibraryItemForMatch(item: BOQItem): RateLibraryItem | null {
    if (!item.linked_rate_id) return null;
    return library.find(l => l.id === item.linked_rate_id) ?? null;
  }

  return (
    <div className="flex flex-col h-full" dir="rtl">
      {/* ── Header ── */}
      <div className="bg-white border-b border-slate-200 px-6 py-4 flex-shrink-0">
        <div className="flex items-center gap-2 mb-3">
          <button onClick={onBack} className="text-blue-600 hover:text-blue-800 text-sm flex items-center gap-1">
            <ChevronRight size={16} />
            <span>المشاريع</span>
          </button>
          <span className="text-slate-300">/</span>
          <span className="text-sm text-slate-600 font-medium truncate max-w-xs">{boqFile.name}</span>
        </div>

        {/* Stats row */}
        <div className="flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-4">
            <div className="text-center">
              <p className="text-2xl font-bold text-slate-900">{items.length}</p>
              <p className="text-xs text-slate-500">إجمالي البنود</p>
            </div>
            <div className="h-8 w-px bg-slate-200" />
            <div className="text-center">
              <p className="text-2xl font-bold text-emerald-600">{pricedPct}%</p>
              <p className="text-xs text-slate-500">نسبة التسعير</p>
            </div>
            {pendingCount > 0 && (
              <>
                <div className="h-8 w-px bg-slate-200" />
                <div className="text-center cursor-pointer" onClick={() => setFilterStatus('pending')}>
                  <p className="text-2xl font-bold text-amber-500">{pendingCount}</p>
                  <p className="text-xs text-slate-500">في الانتظار</p>
                </div>
              </>
            )}
            {unpricedCount > 0 && (
              <>
                <div className="h-8 w-px bg-slate-200" />
                <div className="text-center cursor-pointer" onClick={() => setFilterStatus('unpriced')}>
                  <p className="text-2xl font-bold text-red-500">{unpricedCount}</p>
                  <p className="text-xs text-slate-500">غير مسعَّر</p>
                </div>
              </>
            )}
            <div className="h-8 w-px bg-slate-200" />
            <div className="text-center">
              <p className="text-lg font-bold text-blue-700 ltr">{sarFormat(totalSAR)}</p>
              <p className="text-xs text-slate-500">الإجمالي الكلي</p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={handleBulkPrice}
              disabled={pricingState.active || needsPricingCount === 0}
              title={needsPricingCount === 0 && pendingCount > 0 ? 'كل البنود لديها سعر — راجع البنود في الانتظار للاعتماد' : ''}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-medium rounded-lg transition-colors"
            >
              <RefreshCw size={16} className={pricingState.active ? 'animate-spin' : ''} />
              <span>
                {pricingState.active
                  ? 'جاري التسعير...'
                  : needsPricingCount > 0
                  ? `تسعير جماعي (${needsPricingCount} بند)`
                  : 'تسعير جماعي'}
              </span>
            </button>
            {pendingCount > 0 && (
              <button
                onClick={handleApprovePending}
                disabled={pricingState.active}
                className="flex items-center gap-2 px-4 py-2 bg-amber-500 hover:bg-amber-600 disabled:bg-amber-300 text-white text-sm font-medium rounded-lg transition-colors"
                title="اعتماد كل البنود في الانتظار التي لديها سعر"
              >
                <ThumbsUp size={16} />
                <span>اعتماد ({pendingCount})</span>
              </button>
            )}
            {unpricedCount > 0 && (
              <button
                onClick={handleExportUnpriced}
                disabled={exportingUnpriced}
                className="flex items-center gap-2 px-4 py-2 bg-red-600 hover:bg-red-700 disabled:bg-red-300 text-white text-sm font-medium rounded-lg transition-colors"
                title="تصدير البنود غير المسعّرة لإضافتها لمكتبة الأسعار"
              >
                <FileDown size={16} />
                <span>{exportingUnpriced ? 'جاري التصدير...' : `تصدير غير مسعَّر (${unpricedCount})`}</span>
              </button>
            )}
            <button
              onClick={handleExport}
              disabled={exporting}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-300 text-white text-sm font-medium rounded-lg transition-colors"
            >
              <Download size={16} />
              <span>{exporting ? 'جاري التصدير...' : 'تصدير Excel'}</span>
            </button>
          </div>
        </div>

        {/* Pricing progress bar */}
        {pricingState.active && pricingState.progress && (
          <PricingProgressBar progress={pricingState.progress} />
        )}

        {/* Pricing done banner */}
        {pricingState.done && pricingState.result && (
          <div className="mt-3 flex items-center gap-3 p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-sm">
            <CheckCircle size={16} className="text-emerald-600 flex-shrink-0" />
            <div className="flex items-center gap-4 flex-1 flex-wrap">
              <span className="text-emerald-800 font-medium">اكتمل التسعير</span>
              <span className="text-emerald-700"><span className="font-semibold">{pricingState.result.priced}</span> بند مُسعَّر</span>
              {pricingState.result.failed > 0 && (
                <span className="text-amber-700">
                  <span className="font-semibold">{pricingState.result.failed}</span> بند لم يُطابَق — يمكنك إدخال سعر يدوي بالضغط على
                  <Edit2 size={12} className="inline mx-1" />
                </span>
              )}
              <span className="font-bold text-blue-700 ltr mr-auto">{sarFormat(totalSAR)}</span>
            </div>
            <button onClick={() => setPricingState({ active: false, progress: null, done: false, result: null, error: null })} className="text-slate-400 hover:text-slate-600">
              <XCircle size={16} />
            </button>
          </div>
        )}

        {/* Pricing error banner */}
        {pricingState.error && (
          <div className="mt-3 flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">
            <AlertTriangle size={16} className="flex-shrink-0" />
            <span>خطأ في التسعير: {pricingState.error}</span>
            <button onClick={() => setPricingState(p => ({ ...p, error: null }))} className="mr-auto text-slate-400 hover:text-slate-600"><XCircle size={14} /></button>
          </div>
        )}

        {/* Pending items explanation banner */}
        {!pricingState.active && !pricingState.done && pendingCount > 0 && needsPricingCount === 0 && (
          <div className="mt-3 flex items-start gap-2 p-3 bg-amber-50 border border-amber-200 rounded-lg text-sm">
            <HelpCircle size={16} className="text-amber-600 flex-shrink-0 mt-0.5" />
            <div className="text-amber-800 flex-1">
              <span className="font-semibold">{pendingCount} بند في الانتظار</span> — هذه البنود لديها سعر تلقائي لكن نسبة الثقة أقل من 75%.
              يمكنك <strong>اعتمادها دفعة واحدة</strong> بزر &quot;اعتماد&quot; أعلاه، أو مراجعة وتعديل كل بند منفرداً.
            </div>
          </div>
        )}

        {/* Unpriced items banner */}
        {!pricingState.active && !pricingState.done && unpricedCount > 0 && (
          <div className="mt-3 flex items-start gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-sm">
            <AlertTriangle size={16} className="text-red-500 flex-shrink-0 mt-0.5" />
            <div className="text-red-800 flex-1">
              <span className="font-semibold">{unpricedCount} بند غير مسعَّر</span> — لم يجد النظام تطابقاً في مكتبة الأسعار.
              يمكنك: إدخال سعر يدوي <Edit2 size={12} className="inline mx-0.5" />، أو تصدير هذه البنود لإضافتها للمكتبة ثم إعادة التسعير الجماعي.
            </div>
          </div>
        )}

        {exportError && (
          <div className="mt-3 flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">
            <AlertTriangle size={16} className="flex-shrink-0" />
            <span>{exportError}</span>
          </div>
        )}

        {/* Filter row */}
        <div className="flex items-center gap-3 mt-4 flex-wrap">
          <input
            type="text"
            placeholder="بحث في الوصف أو رقم البند..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            className="flex-1 min-w-48 px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <div className="flex items-center gap-1 flex-wrap">
            {FILTER_OPTIONS.map(opt => {
              const isUnpriced = opt.value === 'unpriced';
              return (
                <button
                  key={opt.value}
                  onClick={() => setFilterStatus(opt.value)}
                  className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors ${
                    filterStatus === opt.value
                      ? isUnpriced ? 'bg-red-600 text-white' : 'bg-blue-600 text-white'
                      : isUnpriced
                        ? 'bg-red-50 text-red-600 hover:bg-red-100'
                        : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                  }`}
                >
                  {opt.label}
                  {isUnpriced && unpricedCount > 0 && (
                    <span className={`mr-1 px-1.5 py-0.5 rounded-full text-xs font-bold ${filterStatus === 'unpriced' ? 'bg-red-500 text-white' : 'bg-red-500 text-white'}`}>
                      {unpricedCount}
                    </span>
                  )}
                </button>
              );
            })}

            <div className="h-5 w-px bg-slate-200 mx-1" />

            {/* Sort buttons */}
            <button
              onClick={() => setSortMode(prev => prev === 'unit_rate_desc' ? 'default' : 'unit_rate_desc')}
              className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors flex items-center gap-1 ${
                sortMode === 'unit_rate_desc'
                  ? 'bg-teal-600 text-white'
                  : 'bg-teal-50 text-teal-700 hover:bg-teal-100'
              }`}
              title="ترتيب حسب أعلى سعر وحدة"
            >
              ↓ أعلى سعر وحدة
            </button>
            <button
              onClick={() => setSortMode(prev => prev === 'total_price_desc' ? 'default' : 'total_price_desc')}
              className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors flex items-center gap-1 ${
                sortMode === 'total_price_desc'
                  ? 'bg-teal-600 text-white'
                  : 'bg-teal-50 text-teal-700 hover:bg-teal-100'
              }`}
              title="ترتيب حسب أعلى إجمالي"
            >
              ↓ أعلى إجمالي
            </button>

            <div className="h-5 w-px bg-slate-200 mx-1" />

            {/* Hide descriptive toggle */}
            <button
              onClick={() => setHideDescriptive(prev => !prev)}
              className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors ${
                hideDescriptive
                  ? 'bg-slate-700 text-white'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
              title="إخفاء البنود الوصفية لتقليل المساحة"
            >
              {hideDescriptive ? '◉ إخفاء الوصفية' : '○ إخفاء الوصفية'}
            </button>
          </div>
        </div>
      </div>

      {/* ── Table ── */}
      <div className="flex-1 overflow-auto">
        <table className="w-full text-sm rtl-table min-w-[1100px]">
          <thead className="bg-slate-800 sticky top-0 z-10">
            <tr>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-16 text-center">رقم القسم</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-28 text-center">كود البند</th>
              <th className="px-3 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 text-right">وصف البند</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-24 text-center">الفئة</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-16 text-center">الوحدة</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-24 text-center">الكمية</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-32 text-left">سعر الوحدة</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-36 text-left">السعر الإجمالي</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-20 text-center">الثقة</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-24 text-center">الحالة</th>
              <th className="px-2 py-3 text-xs font-semibold text-slate-200 border-b border-slate-700 w-24 text-center">إجراءات</th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-slate-100">
            {loading ? (
              Array.from({ length: 8 }).map((_, i) => <SkeletonRow key={i} />)
            ) : loadError ? (
              <tr>
                <td colSpan={11} className="px-6 py-12 text-center">
                  <AlertTriangle size={36} className="mx-auto text-red-400 mb-3" />
                  <p className="text-sm text-red-600 font-medium">{loadError}</p>
                  <button onClick={loadItems} className="mt-3 px-4 py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700">
                    إعادة المحاولة
                  </button>
                </td>
              </tr>
            ) : filteredItems.length === 0 && items.length === 0 ? (
              <tr>
                <td colSpan={11} className="px-6 py-12 text-center">
                  <AlertTriangle size={36} className="mx-auto text-amber-400 mb-3" />
                  <p className="text-sm text-slate-600 font-medium">لا توجد بنود في هذا الملف</p>
                </td>
              </tr>
            ) : filteredItems.length === 0 ? (
              <tr>
                <td colSpan={11} className="px-6 py-12 text-center">
                  <CheckSquare size={36} className="mx-auto text-slate-300 mb-3" />
                  <p className="text-sm text-slate-500">لا توجد بنود تطابق معايير البحث</p>
                </td>
              </tr>
            ) : (
              filteredItems.map((item, idx) => {
                const isDescriptive = item.status === 'descriptive';
                const isManual = item.override_type === 'manual';
                const isRepricingThis = repricingItemId === item.id;
                const isApprovingThis = approvingItemId === item.id;
                const isPending = item.status === 'pending' && item.unit_rate != null && item.unit_rate > 0;
                const divisionCode = item.item_no ? item.item_no.split('-')[0] : '';
                const isUnpricedRow = !isDescriptive && (item.unit_rate == null || item.unit_rate === 0) && (item.quantity ?? 0) > 0;

                return (
                  <tr
                    key={item.id}
                    className={`transition-colors ${
                      isDescriptive
                        ? 'bg-slate-100 font-medium'
                        : isUnpricedRow
                        ? 'bg-red-50/60 hover:bg-red-50'
                        : idx % 2 === 0 ? 'bg-white hover:bg-blue-50/30' : 'bg-slate-50/40 hover:bg-blue-50/30'
                    }`}
                  >
                    <td className="px-2 py-2.5 text-xs text-slate-500 text-center font-mono">{divisionCode || '—'}</td>
                    <td className="px-2 py-2.5 text-xs text-slate-700 text-center font-mono ltr">{item.item_no || '—'}</td>
                    <td className="px-3 py-2.5">
                      <p className={`text-sm leading-snug ${isDescriptive ? 'font-semibold text-slate-800' : 'text-slate-700'}`}>
                        {item.description}
                      </p>
                    </td>
                    <td className="px-2 py-2.5 text-center">
                      {!isDescriptive && item.category ? (
                        <span className="text-xs px-1.5 py-0.5 rounded-md bg-slate-100 text-slate-600 font-medium whitespace-nowrap">
                          {CATEGORY_LABELS_AR[item.category] ?? item.category}
                        </span>
                      ) : (
                        <span className="text-xs text-slate-300">—</span>
                      )}
                    </td>
                    <td className="px-2 py-2.5 text-xs text-slate-600 text-center">{item.unit || '—'}</td>
                    <td className="px-2 py-2.5 text-xs text-slate-700 text-center ltr">
                      {item.quantity != null ? item.quantity.toLocaleString('ar-SA') : '—'}
                    </td>
                    <td className="px-2 py-2.5 text-xs font-medium text-slate-800 ltr text-left">
                      {sarFormat(item.unit_rate)}
                    </td>
                    <td className="px-2 py-2.5 text-xs font-semibold text-slate-900 ltr text-left">
                      {sarFormat(item.total_price)}
                    </td>
                    <td className="px-2 py-2.5 text-center">
                      {!isDescriptive && item.confidence > 0 ? (
                        <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${confidenceColor(item.confidence)}`}>
                          {item.confidence}%
                        </span>
                      ) : (
                        <span className="text-xs text-slate-300">—</span>
                      )}
                    </td>
                    <td className="px-2 py-2.5 text-center">
                      <StatusBadge status={item.status} size="sm" />
                    </td>
                    <td className="px-2 py-2.5">
                      <div className="flex items-center gap-1 justify-center">
                        {isManual ? (
                          <span title="مقفل يدوياً"><Lock size={14} className="text-blue-400" /></span>
                        ) : (
                          !isDescriptive && (
                            <button
                              onClick={() => handleRepriceSingle(item.id)}
                              disabled={isRepricingThis || pricingState.active}
                              title="إعادة التسعير"
                              className="w-7 h-7 rounded-md flex items-center justify-center text-slate-400 hover:bg-amber-50 hover:text-amber-600 transition-colors disabled:opacity-40"
                            >
                              <RefreshCw size={14} className={isRepricingThis ? 'animate-spin' : ''} />
                            </button>
                          )
                        )}
                        {isPending && (
                          <button
                            onClick={() => handleApproveSingle(item.id)}
                            disabled={isApprovingThis}
                            title="اعتماد هذا البند"
                            className="w-7 h-7 rounded-md flex items-center justify-center text-emerald-500 hover:bg-emerald-50 hover:text-emerald-700 transition-colors disabled:opacity-40"
                          >
                            <Check size={14} className={isApprovingThis ? 'animate-pulse' : ''} />
                          </button>
                        )}
                        {!isDescriptive && item.linked_rate_id && (
                          <button
                            onClick={() => setMatchItemState(item)}
                            title="عرض التطابق"
                            className="w-7 h-7 rounded-md flex items-center justify-center text-slate-400 hover:bg-blue-50 hover:text-blue-600 transition-colors"
                          >
                            <Eye size={14} />
                          </button>
                        )}
                        {!isDescriptive && (
                          <button
                            onClick={() => setEditItem(item)}
                            title="تعديل السعر"
                            className={`w-7 h-7 rounded-md flex items-center justify-center transition-colors ${
                              isUnpricedRow
                                ? 'text-red-400 hover:bg-red-50 hover:text-red-600'
                                : 'text-slate-400 hover:bg-emerald-50 hover:text-emerald-600'
                            }`}
                          >
                            <Edit2 size={14} />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {editItem && (
        <PriceBreakdownModal
          item={editItem}
          onClose={() => setEditItem(null)}
          onSaved={async () => {
            setEditItem(null);
            await refreshItems();
          }}
        />
      )}

      {matchItemState && (
        <MatchDetailsModal
          item={matchItemState}
          libraryItem={getLibraryItemForMatch(matchItemState)}
          onClose={() => setMatchItemState(null)}
        />
      )}
    </div>
  );
}

function SkeletonRow() {
  return (
    <tr>
      {Array.from({ length: 11 }).map((_, i) => (
        <td key={i} className="px-2 py-3">
          <div className="skeleton h-4 rounded w-full" />
        </td>
      ))}
    </tr>
  );
}

function PricingProgressBar({ progress }: { progress: PricingProgress }) {
  const pct = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;

  return (
    <div className="mt-3 p-4 bg-blue-50 border border-blue-200 rounded-xl">
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-2">
          <RefreshCw size={14} className="text-blue-600 animate-spin" />
          <span className="text-sm font-semibold text-blue-800">جاري التسعير بند بند...</span>
        </div>
        <div className="flex items-center gap-4">
          <span className="text-xs text-blue-700"><span className="font-bold">{progress.pricedSoFar}</span> مُسعَّر</span>
          {progress.failedSoFar > 0 && (
            <span className="text-xs text-amber-700"><span className="font-bold">{progress.failedSoFar}</span> بدون تطابق</span>
          )}
          <span className="text-sm font-bold text-blue-900 ltr">
            {new Intl.NumberFormat('ar-SA', { style: 'currency', currency: 'SAR', maximumFractionDigits: 0 }).format(progress.runningTotal)}
          </span>
        </div>
      </div>

      <div className="h-2.5 bg-blue-100 rounded-full overflow-hidden mb-2">
        <div
          className="h-full bg-gradient-to-r from-blue-500 to-emerald-500 rounded-full transition-all duration-300"
          style={{ width: `${Math.max(pct, 2)}%` }}
        />
      </div>

      <div className="flex items-center justify-between">
        <p className="text-xs text-blue-700 truncate max-w-md" title={progress.currentItem}>
          {progress.currentItem}
        </p>
        <span className="text-xs font-semibold text-blue-800 flex-shrink-0 mr-2">
          {progress.current} / {progress.total} ({pct}%)
        </span>
      </div>
    </div>
  );
}
