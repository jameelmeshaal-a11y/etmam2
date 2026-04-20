import { useState, useEffect } from 'react';
import { X, Save, Cpu, Zap, Droplets, ChevronDown, ChevronUp } from 'lucide-react';
import { supabase } from '../lib/supabase';
import type { BOQItem } from '../types';

interface BMSCalculatorModalProps {
  item: BOQItem;
  boqFileId: string;
  onClose: () => void;
  onSaved: () => void;
}

interface Equipment {
  id: string;
  name: string;
  ai: number;
  di: number;
  ao: number;
  do_: number;
}

interface EquipmentGroup {
  id: string;
  label: string;
  color: string;
  icon: React.ReactNode;
  items: Equipment[];
}

const EQUIPMENT_GROUPS: EquipmentGroup[] = [
  {
    id: 'hvac',
    label: 'أنظمة HVAC الميكانيكية',
    color: 'blue',
    icon: null,
    items: [
      { id: 'ahu',           name: 'وحدة معالجة هواء AHU',       ai: 3, di: 3, ao: 2, do_: 2 },
      { id: 'fcu',           name: 'وحدة فانكويل FCU',            ai: 2, di: 2, ao: 1, do_: 1 },
      { id: 'chiller',       name: 'مبرد مياه Chiller',           ai: 5, di: 4, ao: 2, do_: 3 },
      { id: 'pump_chw',      name: 'مضخة مياه باردة/ساخنة',      ai: 1, di: 2, ao: 1, do_: 1 },
      { id: 'cooling_tower', name: 'برج تبريد Cooling Tower',     ai: 2, di: 3, ao: 1, do_: 2 },
      { id: 'vav',           name: 'صندوق VAV',                   ai: 1, di: 1, ao: 1, do_: 1 },
      { id: 'pac',           name: 'وحدة تبريد PAC/Split',        ai: 1, di: 2, ao: 0, do_: 1 },
      { id: 'exhaust_fan',   name: 'مروحة طرد هواء',              ai: 0, di: 2, ao: 0, do_: 1 },
      { id: 'damper',        name: 'شيش هواء (Damper)',            ai: 0, di: 1, ao: 1, do_: 1 },
    ],
  },
  {
    id: 'electrical',
    label: 'أنظمة الكهرباء',
    color: 'amber',
    icon: null,
    items: [
      { id: 'mdb',       name: 'لوحة رئيسية MDB',             ai: 3, di: 3, ao: 0, do_: 2 },
      { id: 'ups',       name: 'منظومة UPS',                  ai: 2, di: 2, ao: 0, do_: 2 },
      { id: 'generator', name: 'مولد كهربائي Generator',      ai: 3, di: 4, ao: 0, do_: 3 },
      { id: 'lighting',  name: 'نظام إضاءة (لوحة)',           ai: 0, di: 2, ao: 0, do_: 1 },
      { id: 'sdb',       name: 'لوحة فرعية SDB',              ai: 1, di: 2, ao: 0, do_: 1 },
    ],
  },
  {
    id: 'plumbing',
    label: 'السباكة والصرف الصحي',
    color: 'teal',
    icon: null,
    items: [
      { id: 'water_tank',    name: 'خزان مياه (علوي/سفلي)',   ai: 1, di: 2, ao: 0, do_: 1 },
      { id: 'booster_pump',  name: 'مضخة ضغط / رفع مياه',     ai: 1, di: 1, ao: 1, do_: 1 },
      { id: 'sewage_pump',   name: 'مضخة صرف صحي',            ai: 0, di: 2, ao: 0, do_: 1 },
      { id: 'flow_meter',    name: 'عداد تدفق مياه',           ai: 1, di: 1, ao: 0, do_: 0 },
    ],
  },
];

const sarFormat = (v: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR', maximumFractionDigits: 0 }).format(v);

function pointsPerEquipment(eq: Equipment): number {
  return eq.ai + eq.di + eq.ao + eq.do_;
}

const COLOR_MAP: Record<string, { dot: string; header: string; badge: string; row: string }> = {
  blue: {
    dot: 'bg-blue-500',
    header: 'bg-blue-50 border-blue-100',
    badge: 'bg-blue-100 text-blue-700',
    row: 'hover:bg-blue-50/50',
  },
  amber: {
    dot: 'bg-amber-500',
    header: 'bg-amber-50 border-amber-100',
    badge: 'bg-amber-100 text-amber-700',
    row: 'hover:bg-amber-50/50',
  },
  teal: {
    dot: 'bg-teal-500',
    header: 'bg-teal-50 border-teal-100',
    badge: 'bg-teal-100 text-teal-700',
    row: 'hover:bg-teal-50/50',
  },
};

const GROUP_ICONS: Record<string, React.ReactNode> = {
  hvac:       <Cpu size={16} className="text-blue-600" />,
  electrical: <Zap size={16} className="text-amber-600" />,
  plumbing:   <Droplets size={16} className="text-teal-600" />,
};

type Counts = Record<string, number>;

function isBMSItem(description: string): boolean {
  const d = description.trim();
  const dl = d.toLowerCase();

  // Exclude "لزوم نظام BMS" suffix patterns (item FOR bms, not IS bms)
  if (/لزوم\s+نظام\s+bms\s*$/i.test(dl)) return false;
  if (/لزوم\s+.*bms\s*$/i.test(dl)) return false;

  // Must be primary subject: BMS/نظام إدارة المباني as the main item
  if (/^[^،,\n]*\(\s*BMS\s*\)/i.test(d)) return true;   // "نظام إدارة المباني ( BMS )"
  if (/^[^،,\n]*\(\s*BAS\s*\)/i.test(d)) return true;
  if (/^\s*(نظام\s+إدارة\s+المباني|نظام\s+ادارة\s+المباني)/i.test(d)) return true;
  if (/^\s*bms\b/i.test(dl)) return true;
  if (/^\s*bas\b/i.test(dl)) return true;
  if (/^\s*(building\s+management|building\s+automation)/i.test(dl)) return true;
  if (/^\s*(نظام\s+)?(ddc|scada)\b/i.test(dl)) return true;

  // Multi-line: long description with BMS in inner section header
  if (d.length > 200 && /\n[^\n]*\(\s*BMS\s*\)/i.test(d)) return true;

  return false;
}

export { isBMSItem };

export default function BMSCalculatorModal({ item, boqFileId, onClose, onSaved }: BMSCalculatorModalProps) {
  const [counts, setCounts] = useState<Counts>({});
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  const [loading, setLoading] = useState(true);

  // Load existing calculation if any
  useEffect(() => {
    async function loadExisting() {
      const { data } = await supabase
        .from('bms_calculations')
        .select('equipment_data')
        .eq('boq_item_id', item.id)
        .maybeSingle();
      if (data?.equipment_data) {
        setCounts(data.equipment_data as Counts);
      }
      setLoading(false);
    }
    loadExisting();
  }, [item.id]);

  function setCount(id: string, val: number) {
    setCounts(prev => ({ ...prev, [id]: Math.max(0, val) }));
  }

  // Totals
  const allEquipment = EQUIPMENT_GROUPS.flatMap(g => g.items);
  let totalAI = 0, totalDI = 0, totalAO = 0, totalDO = 0;
  for (const eq of allEquipment) {
    const n = counts[eq.id] ?? 0;
    totalAI += eq.ai * n;
    totalDI += eq.di * n;
    totalAO += eq.ao * n;
    totalDO += eq.do_ * n;
  }
  const totalPoints = totalAI + totalDI + totalAO + totalDO;
  const totalCost = totalPoints * 500;

  async function handleSave() {
    if (totalPoints === 0) {
      setError('يجب إدخال عدد المعدات للحصول على التكلفة');
      return;
    }
    setSaving(true);
    setError(null);

    try {
      const { data: { user } } = await supabase.auth.getUser();

      // Upsert the calculation record
      const { error: calcErr } = await supabase
        .from('bms_calculations')
        .upsert(
          {
            boq_item_id: item.id,
            boq_file_id: boqFileId,
            equipment_data: counts,
            total_points: totalPoints,
            total_cost: totalCost,
            price_per_point: 500,
            created_by: user?.id,
            updated_at: new Date().toISOString(),
          },
          { onConflict: 'boq_item_id' }
        );

      if (calcErr) throw calcErr;

      // Apply price to the BOQ item — use save_manual_price RPC
      const unitRate = item.quantity && item.quantity > 0
        ? totalCost / item.quantity
        : totalCost;

      const { error: rpcErr } = await supabase.rpc('save_manual_price', {
        p_boq_item_id: item.id,
        p_unit_rate: unitRate,
        p_materials: 0,
        p_labor: 0,
        p_equipment: 0,
        p_logistics: 0,
        p_risk: 0,
        p_profit: 0,
        p_linked_rate_id: null,
      });

      if (rpcErr) {
        // Fallback direct update
        await supabase
          .from('boq_items')
          .update({
            unit_rate: unitRate,
            total_price: totalCost,
            status: 'manual',
            override_type: 'manual',
            updated_at: new Date().toISOString(),
          })
          .eq('id', item.id);
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
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col fade-in">

        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 flex-shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-slate-800 flex items-center justify-center">
              <Cpu size={18} className="text-white" />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900">حاسبة نقاط BMS</h2>
              <p className="text-xs text-slate-500 mt-0.5">سعر النقطة: 500 ريال ثابت</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600"
          >
            <X size={18} />
          </button>
        </div>

        {/* Item info */}
        <div className="px-6 py-3 bg-slate-50 border-b border-slate-100 flex-shrink-0">
          <p className="text-xs text-slate-500 mb-1">بند نظام إدارة المباني</p>
          <p className="text-sm font-medium text-slate-800 leading-snug line-clamp-2">{item.description}</p>
          {item.quantity != null && item.quantity > 0 && (
            <p className="text-xs text-slate-500 mt-1">الكمية: <span className="font-medium text-slate-700">{item.quantity} {item.unit}</span></p>
          )}
        </div>

        {/* Stats bar */}
        <div className="grid grid-cols-5 divide-x divide-x-reverse divide-slate-100 bg-white border-b border-slate-100 flex-shrink-0">
          {[
            { label: 'AI', value: totalAI, color: 'text-blue-600' },
            { label: 'DI', value: totalDI, color: 'text-slate-600' },
            { label: 'AO', value: totalAO, color: 'text-emerald-600' },
            { label: 'DO', value: totalDO, color: 'text-rose-600' },
            { label: 'إجمالي النقاط', value: totalPoints, color: 'text-slate-900' },
          ].map(s => (
            <div key={s.label} className="py-3 text-center">
              <p className={`text-lg font-bold ${s.color}`}>{s.value}</p>
              <p className="text-xs text-slate-400 mt-0.5">{s.label}</p>
            </div>
          ))}
        </div>

        {/* Equipment list */}
        <div className="flex-1 overflow-y-auto">
          {loading ? (
            <div className="py-12 text-center text-slate-400 text-sm">جاري التحميل...</div>
          ) : (
            EQUIPMENT_GROUPS.map(group => {
              const c = COLOR_MAP[group.color];
              const isCollapsed = collapsed[group.id];
              const groupPoints = group.items.reduce((sum, eq) => {
                const n = counts[eq.id] ?? 0;
                return sum + pointsPerEquipment(eq) * n;
              }, 0);

              return (
                <div key={group.id} className="border-b border-slate-100">
                  <button
                    className={`w-full flex items-center justify-between px-6 py-3 border-b ${c.header} transition-colors`}
                    onClick={() => setCollapsed(prev => ({ ...prev, [group.id]: !isCollapsed }))}
                  >
                    <div className="flex items-center gap-2">
                      <span className={`w-2.5 h-2.5 rounded-full ${c.dot} flex-shrink-0`} />
                      <span className="text-sm font-semibold text-slate-800">{group.label}</span>
                      {GROUP_ICONS[group.id]}
                    </div>
                    <div className="flex items-center gap-3">
                      {groupPoints > 0 && (
                        <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${c.badge}`}>
                          {groupPoints} نقطة
                        </span>
                      )}
                      {isCollapsed ? <ChevronDown size={16} className="text-slate-400" /> : <ChevronUp size={16} className="text-slate-400" />}
                    </div>
                  </button>

                  {!isCollapsed && (
                    <div>
                      {group.items.map(eq => {
                        const n = counts[eq.id] ?? 0;
                        const pts = pointsPerEquipment(eq);

                        return (
                          <div
                            key={eq.id}
                            className={`flex items-center gap-3 px-6 py-3 border-b border-slate-50 transition-colors ${c.row}`}
                          >
                            {/* Name */}
                            <div className="flex-1 min-w-0">
                              <p className="text-sm text-slate-800 leading-tight">{eq.name}</p>
                              <div className="flex items-center gap-1.5 mt-1">
                                <PointBadge label="AI" value={eq.ai} color="bg-blue-50 text-blue-600" />
                                <PointBadge label="DI" value={eq.di} color="bg-slate-100 text-slate-600" />
                                <PointBadge label="AO" value={eq.ao} color="bg-emerald-50 text-emerald-600" />
                                <PointBadge label="DO" value={eq.do_} color="bg-rose-50 text-rose-600" />
                                <span className="text-xs text-slate-400 mr-1">{pts} نقطة/وحدة</span>
                              </div>
                            </div>

                            {/* Counter */}
                            <div className="flex items-center gap-2 flex-shrink-0">
                              {n > 0 && (
                                <span className="text-xs text-slate-500 w-20 text-left ltr">
                                  = {(pts * n).toLocaleString('en-US')} نقطة
                                </span>
                              )}
                              <div className="flex items-center gap-1">
                                <button
                                  onClick={() => setCount(eq.id, n - 1)}
                                  className="w-7 h-7 rounded-lg border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-100 hover:border-slate-300 transition-colors font-bold"
                                >
                                  −
                                </button>
                                <span className="w-8 text-center text-sm font-semibold text-slate-800">{n}</span>
                                <button
                                  onClick={() => setCount(eq.id, n + 1)}
                                  className="w-7 h-7 rounded-lg border border-slate-200 flex items-center justify-center text-slate-500 hover:bg-slate-100 hover:border-slate-300 transition-colors font-bold"
                                >
                                  +
                                </button>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })
          )}
        </div>

        {/* Footer */}
        <div className="flex-shrink-0 border-t border-slate-100 bg-white">
          {/* Cost summary */}
          <div className="px-6 py-4 flex items-center justify-between bg-slate-50 border-b border-slate-100">
            <div>
              <p className="text-xs text-slate-500">إجمالي نقاط النظام</p>
              <p className="text-xl font-bold text-slate-900">{totalPoints.toLocaleString('en-US')} نقطة</p>
            </div>
            <div className="text-left">
              <p className="text-xs text-slate-500">التكلفة التقديرية (500 ر.س / نقطة)</p>
              <p className="text-xl font-bold text-blue-700 ltr">{sarFormat(totalCost)}</p>
            </div>
          </div>

          {error && (
            <div className="px-6 py-2">
              <p className="text-sm text-red-600 bg-red-50 px-3 py-2 rounded-lg">{error}</p>
            </div>
          )}

          <div className="flex items-center gap-3 px-6 py-4">
            <button
              onClick={handleSave}
              disabled={saving || totalPoints === 0}
              className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-slate-800 hover:bg-slate-900 disabled:bg-slate-300 text-white text-sm font-medium rounded-xl transition-colors"
            >
              <Save size={16} />
              <span>
                {saving ? 'جاري الحفظ...' : totalPoints > 0 ? `تطبيق السعر ${sarFormat(totalCost)}` : 'أدخل عدد المعدات أولاً'}
              </span>
            </button>
            <button
              onClick={onClose}
              className="px-4 py-2.5 text-sm text-slate-600 hover:bg-slate-100 rounded-xl transition-colors"
            >
              إلغاء
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PointBadge({ label, value, color }: { label: string; value: number; color: string }) {
  if (value === 0) return <span className={`text-xs px-1.5 py-0.5 rounded font-medium opacity-30 ${color}`}>{value}{label}</span>;
  return <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${color}`}>{value}{label}</span>;
}
