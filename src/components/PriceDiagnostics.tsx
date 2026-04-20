import { RefreshCw, AlertTriangle, BarChart3 } from 'lucide-react';
import { repriceSingleItem } from '../lib/pricingEngine';
import { getAllLibraryItems } from '../lib/priceLibrary';
import { useState } from 'react';
import type { BOQItem } from '../types';
import StatusBadge from './StatusBadge';

interface PriceDiagnosticsProps {
  boqFileId: string;
  items: BOQItem[];
  onRepriced?: () => void;
}

interface ConfidenceBucket {
  label: string;
  min: number;
  max: number;
  color: string;
  bgColor: string;
}

const BUCKETS: ConfidenceBucket[] = [
  { label: '0–25%', min: 0, max: 25, color: 'bg-red-500', bgColor: 'bg-red-50 border-red-200' },
  { label: '25–50%', min: 25, max: 50, color: 'bg-orange-400', bgColor: 'bg-orange-50 border-orange-200' },
  { label: '50–75%', min: 50, max: 75, color: 'bg-amber-400', bgColor: 'bg-amber-50 border-amber-200' },
  { label: '75–100%', min: 75, max: 101, color: 'bg-emerald-500', bgColor: 'bg-emerald-50 border-emerald-200' },
];

const sarFormat = (v: number | null) =>
  v == null ? '—' : new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR' }).format(v);

export default function PriceDiagnostics({ items, onRepriced }: PriceDiagnosticsProps) {
  const [repricingId, setRepricingId] = useState<string | null>(null);

  const nonDescriptive = items.filter(i => i.status !== 'descriptive');
  const withConfidence = nonDescriptive.filter(i => i.confidence > 0);
  const suspicious = items.filter(i => i.confidence > 0 && i.confidence < 75 && i.status !== 'descriptive');
  const needsReview = items.filter(i => i.status === 'needs_review');

  const bucketCounts = BUCKETS.map(b => ({
    ...b,
    count: withConfidence.filter(i => i.confidence >= b.min && i.confidence < b.max).length,
  }));

  const maxCount = Math.max(...bucketCounts.map(b => b.count), 1);

  async function handleReprice(itemId: string) {
    setRepricingId(itemId);
    try {
      const library = await getAllLibraryItems();
      await repriceSingleItem(itemId, library);
      onRepriced?.();
    } catch (e) {
      console.error(e);
    } finally {
      setRepricingId(null);
    }
  }

  return (
    <div className="space-y-6" dir="rtl">
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="flex items-center gap-2 mb-4">
          <BarChart3 size={18} className="text-blue-600" />
          <h3 className="text-sm font-semibold text-slate-900">توزيع درجات الثقة</h3>
          <span className="text-xs text-slate-400 mr-1">({withConfidence.length} بند مُسعَّر)</span>
        </div>

        <div className="space-y-3">
          {bucketCounts.map(bucket => {
            const pct = maxCount > 0 ? (bucket.count / maxCount) * 100 : 0;
            return (
              <div key={bucket.label} className="flex items-center gap-3">
                <span className="w-16 text-xs text-slate-500 text-left flex-shrink-0">{bucket.label}</span>
                <div className="flex-1 h-6 bg-slate-100 rounded-full overflow-hidden">
                  <div
                    className={`h-full rounded-full transition-all duration-500 ${bucket.color}`}
                    style={{ width: `${pct}%` }}
                  />
                </div>
                <span className="w-10 text-xs font-medium text-slate-700 text-right flex-shrink-0">
                  {bucket.count}
                </span>
              </div>
            );
          })}
        </div>

        <div className="mt-4 grid grid-cols-2 gap-3">
          <div className={`p-3 rounded-lg border text-center ${BUCKETS[3].bgColor}`}>
            <p className="text-2xl font-bold text-emerald-700">
              {bucketCounts[3].count}
            </p>
            <p className="text-xs text-emerald-600 mt-0.5">ثقة عالية (75%+)</p>
          </div>
          <div className={`p-3 rounded-lg border ${BUCKETS[0].bgColor} text-center`}>
            <p className="text-2xl font-bold text-red-700">
              {bucketCounts[0].count + bucketCounts[1].count}
            </p>
            <p className="text-xs text-red-600 mt-0.5">تطابق ضعيف (&lt;50%)</p>
          </div>
        </div>
      </div>

      {suspicious.length > 0 && (
        <div className="bg-white rounded-xl border border-amber-200 p-5">
          <div className="flex items-center gap-2 mb-4">
            <AlertTriangle size={18} className="text-amber-600" />
            <h3 className="text-sm font-semibold text-slate-900">تطابقات مشبوهة</h3>
            <span className="text-xs text-amber-600 bg-amber-100 px-2 py-0.5 rounded-full font-medium">{suspicious.length}</span>
          </div>
          <div className="space-y-2 max-h-64 overflow-y-auto">
            {suspicious.map(item => (
              <div key={item.id} className="flex items-center gap-3 p-3 bg-slate-50 rounded-lg border border-slate-100">
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium text-slate-800 truncate">{item.description}</p>
                  <div className="flex items-center gap-2 mt-1">
                    <span className="text-xs text-slate-500">الثقة:</span>
                    <span className="text-xs font-bold text-amber-700">{item.confidence}%</span>
                    <span className="text-xs text-slate-400">|</span>
                    <span className="text-xs text-slate-500">السعر:</span>
                    <span className="text-xs text-slate-700 ltr">{sarFormat(item.unit_rate)}</span>
                  </div>
                </div>
                <button
                  onClick={() => handleReprice(item.id)}
                  disabled={repricingId === item.id}
                  title="إعادة التسعير"
                  className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-blue-50 hover:text-blue-600 transition-colors disabled:opacity-40 flex-shrink-0"
                >
                  <RefreshCw size={14} className={repricingId === item.id ? 'animate-spin' : ''} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {needsReview.length > 0 && (
        <div className="bg-white rounded-xl border border-red-200 p-5">
          <div className="flex items-center gap-2 mb-4">
            <AlertTriangle size={18} className="text-red-600" />
            <h3 className="text-sm font-semibold text-slate-900">يحتاج مراجعة</h3>
            <span className="text-xs text-red-600 bg-red-100 px-2 py-0.5 rounded-full font-medium">{needsReview.length}</span>
          </div>
          <div className="space-y-2 max-h-64 overflow-y-auto">
            {needsReview.map(item => (
              <div key={item.id} className="flex items-center gap-3 p-3 bg-slate-50 rounded-lg border border-slate-100">
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium text-slate-800 truncate">{item.description}</p>
                  <div className="flex items-center gap-2 mt-1">
                    <StatusBadge status={item.status} size="sm" />
                    <span className="text-xs text-slate-500 ltr">{sarFormat(item.unit_rate)}</span>
                  </div>
                </div>
                <button
                  onClick={() => handleReprice(item.id)}
                  disabled={repricingId === item.id}
                  className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-red-50 hover:text-red-600 transition-colors disabled:opacity-40 flex-shrink-0"
                >
                  <RefreshCw size={14} className={repricingId === item.id ? 'animate-spin' : ''} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {suspicious.length === 0 && needsReview.length === 0 && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-6 text-center">
          <p className="text-sm font-medium text-emerald-800">لا توجد مشكلات في التسعير</p>
          <p className="text-xs text-emerald-600 mt-1">جميع البنود ضمن النطاق المقبول</p>
        </div>
      )}
    </div>
  );
}
