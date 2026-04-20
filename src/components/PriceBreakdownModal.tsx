import { useState } from 'react';
import { X, Save, Calculator } from 'lucide-react';
import { supabase } from '../lib/supabase';
import type { BOQItem } from '../types';

interface PriceBreakdownModalProps {
  item: BOQItem;
  onClose: () => void;
  onSaved: () => void;
}

type InputMode = 'unit' | 'total';

const sarFormat = (v: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR' }).format(v);

export default function PriceBreakdownModal({ item, onClose, onSaved }: PriceBreakdownModalProps) {
  const qty = item.quantity ?? 0;

  const [mode, setMode] = useState<InputMode>('unit');
  const [unitRate, setUnitRate] = useState<string>(item.unit_rate ? String(item.unit_rate) : '');
  const [totalInput, setTotalInput] = useState<string>(item.total_price ? String(item.total_price) : '');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const derivedUnitRate = mode === 'unit'
    ? parseFloat(unitRate) || 0
    : qty > 0 ? (parseFloat(totalInput) || 0) / qty : 0;

  const derivedTotal = mode === 'total'
    ? parseFloat(totalInput) || 0
    : derivedUnitRate * qty;

  async function handleSave() {
    if (derivedUnitRate <= 0) {
      setError('يجب إدخال سعر أكبر من صفر');
      return;
    }

    setSaving(true);
    setError(null);

    try {
      const { error: rpcError } = await supabase.rpc('save_manual_price', {
        p_boq_item_id: item.id,
        p_unit_rate: derivedUnitRate,
        p_materials: 0,
        p_labor: 0,
        p_equipment: 0,
        p_logistics: 0,
        p_risk: 0,
        p_profit: 0,
        p_linked_rate_id: null,
      });

      if (rpcError) {
        const { error: updateError } = await supabase
          .from('boq_items')
          .update({
            unit_rate: derivedUnitRate,
            total_price: derivedTotal,
            status: 'manual',
            override_type: 'manual',
            updated_at: new Date().toISOString(),
          })
          .eq('id', item.id);

        if (updateError) throw updateError;
      }

      onSaved();
      onClose();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm" dir="rtl">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-md fade-in">

        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100">
          <div>
            <h2 className="text-base font-semibold text-slate-900">تعديل السعر</h2>
            <p className="text-xs text-slate-500 mt-0.5">إدخال سعر يدوي للبند</p>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600"
          >
            <X size={18} />
          </button>
        </div>

        {/* Item info */}
        <div className="px-6 py-4 bg-slate-50 border-b border-slate-100">
          <p className="text-sm font-medium text-slate-800 leading-relaxed">{item.description}</p>
          <div className="flex items-center gap-4 mt-2">
            <span className="text-xs text-slate-500">
              الوحدة: <span className="font-medium text-blue-700 bg-blue-50 px-2 py-0.5 rounded mr-1">{item.unit || '—'}</span>
            </span>
            <span className="text-xs text-slate-500">
              الكمية: <span className="font-medium text-slate-700 mr-1">{qty.toLocaleString('en-US')}</span>
            </span>
          </div>
        </div>

        {/* Mode toggle */}
        <div className="px-6 pt-5">
          <div className="flex rounded-lg border border-slate-200 overflow-hidden mb-5">
            <button
              onClick={() => { setMode('unit'); setError(null); }}
              className={`flex-1 py-2 text-sm font-medium transition-colors ${
                mode === 'unit'
                  ? 'bg-blue-600 text-white'
                  : 'bg-white text-slate-600 hover:bg-slate-50'
              }`}
            >
              إدخال سعر الوحدة
            </button>
            <button
              onClick={() => { setMode('total'); setError(null); }}
              className={`flex-1 py-2 text-sm font-medium transition-colors border-r border-slate-200 ${
                mode === 'total'
                  ? 'bg-blue-600 text-white'
                  : 'bg-white text-slate-600 hover:bg-slate-50'
              }`}
            >
              إدخال السعر الإجمالي
            </button>
          </div>

          {mode === 'unit' ? (
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">
                سعر الوحدة (ر.س)
              </label>
              <div className="relative">
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={unitRate}
                  onChange={e => { setUnitRate(e.target.value); setError(null); }}
                  placeholder="أدخل سعر الوحدة..."
                  autoFocus
                  className="w-full ltr px-3 py-3 text-base border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-right"
                />
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-xs text-slate-400 font-medium">ر.س</span>
              </div>
              {qty > 0 && derivedUnitRate > 0 && (
                <p className="text-xs text-slate-500 mt-2">
                  السعر الإجمالي: <span className="font-semibold text-slate-700 ltr">{sarFormat(derivedTotal)}</span>
                </p>
              )}
            </div>
          ) : (
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">
                السعر الإجمالي (ر.س)
              </label>
              <div className="relative">
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={totalInput}
                  onChange={e => { setTotalInput(e.target.value); setError(null); }}
                  placeholder="أدخل السعر الإجمالي..."
                  autoFocus
                  className="w-full ltr px-3 py-3 text-base border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-right"
                />
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-xs text-slate-400 font-medium">ر.س</span>
              </div>
              {qty > 0 && derivedUnitRate > 0 && (
                <p className="text-xs text-slate-500 mt-2">
                  سعر الوحدة المحسوب: <span className="font-semibold text-slate-700 ltr">{sarFormat(derivedUnitRate)}</span>
                </p>
              )}
              {qty === 0 && mode === 'total' && (
                <p className="text-xs text-amber-600 mt-2">تنبيه: الكمية صفر — لن يمكن حساب سعر الوحدة</p>
              )}
            </div>
          )}
        </div>

        {/* Summary */}
        {derivedUnitRate > 0 && (
          <div className="mx-6 mt-4 p-3 bg-blue-50 border border-blue-100 rounded-lg">
            <div className="flex items-center gap-2 mb-1">
              <Calculator size={14} className="text-blue-600" />
              <span className="text-xs font-semibold text-blue-800">ملخص</span>
            </div>
            <div className="flex items-center justify-between">
              <span className="text-xs text-blue-700">سعر الوحدة</span>
              <span className="text-sm font-bold text-blue-900 ltr">{sarFormat(derivedUnitRate)}</span>
            </div>
            {qty > 0 && (
              <div className="flex items-center justify-between mt-0.5">
                <span className="text-xs text-blue-700">الإجمالي ({qty.toLocaleString('en-US')} {item.unit})</span>
                <span className="text-sm font-bold text-blue-900 ltr">{sarFormat(derivedTotal)}</span>
              </div>
            )}
          </div>
        )}

        {error && (
          <div className="mx-6 mt-3">
            <p className="text-sm text-red-600 bg-red-50 px-3 py-2 rounded-lg">{error}</p>
          </div>
        )}

        {/* Actions */}
        <div className="flex items-center gap-3 px-6 py-4 mt-2 border-t border-slate-100">
          <button
            onClick={handleSave}
            disabled={saving || derivedUnitRate <= 0}
            className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-medium rounded-lg transition-colors"
          >
            <Save size={16} />
            <span>{saving ? 'جاري الحفظ...' : 'حفظ السعر'}</span>
          </button>
          <button
            onClick={onClose}
            className="px-4 py-2.5 text-sm text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
          >
            إلغاء
          </button>
        </div>
      </div>
    </div>
  );
}
