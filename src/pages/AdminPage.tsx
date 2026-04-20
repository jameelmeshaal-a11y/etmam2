import { useState, useEffect } from 'react';
import { ShieldCheck, Database, Clock, RefreshCw, CheckCircle, AlertTriangle, Activity } from 'lucide-react';
import { supabase } from '../lib/supabase';

interface AdminPageProps {
  onNavigate?: (page: string, data?: unknown) => void;
}

interface Migration {
  version: string;
  name: string;
  inserted_at: string;
}

interface AuditEntry {
  id: string;
  boq_item_id: string;
  unit_rate: number;
  source_type: string;
  override_type: string | null;
  created_at: string;
  created_by: string;
}

interface GovernanceWall {
  number: number;
  name: string;
  description: string;
  active: boolean;
}

const GOVERNANCE_WALLS: GovernanceWall[] = [
  { number: 1, name: 'حماية الإدخال اليدوي', description: 'يمنع إعادة تسعير البنود المقفولة يدوياً', active: true },
  { number: 2, name: 'قفل البند اليدوي', description: 'البنود التي تم تسعيرها يدوياً لا تتأثر بالتسعير التلقائي', active: true },
  { number: 3, name: 'بوابة الكلمات المفتاحية', description: 'تصفية النتائج بناءً على مطابقة الكلمات المفتاحية', active: true },
  { number: 4, name: 'حدود الحوكمة', description: 'لا يُقبل أي سعر خارج النطاق المسموح به في المكتبة', active: true },
  { number: 5, name: 'التحقق من التصدير', description: 'يرفض التصدير إذا تجاوز الفارق ±2.5%', active: true },
];

const sarFormat = (v: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR' }).format(v);

export default function AdminPage({ onNavigate: _onNavigate }: AdminPageProps) {
  const [migrations, setMigrations] = useState<Migration[]>([]);
  const [auditLog, setAuditLog] = useState<AuditEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [cleanupRunning, setCleanupRunning] = useState(false);
  const [cleanupResult, setCleanupResult] = useState<string | null>(null);

  useEffect(() => {
    async function load() {
      setLoading(true);
      try {
        const migrationsRpc = await supabase.rpc('list_migrations');
        const migrationsRes = migrationsRpc.data ?? [];
        const auditRes = await supabase.from('rate_sources').select('*').order('created_at', { ascending: false }).limit(20);

        setMigrations(migrationsRes as Migration[]);
        setAuditLog((auditRes.data ?? []) as AuditEntry[]);
      } catch (e) {
        console.error(e);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, []);

  async function handleStaleCleanup() {
    setCleanupRunning(true);
    setCleanupResult(null);
    try {
      const { error, count } = await supabase
        .from('boq_items')
        .update({ status: 'stale_price' })
        .eq('status', 'approved')
        .lt('updated_at', new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString())
        .neq('override_type', 'manual');

      if (error) throw error;
      setCleanupResult(`تم تحديث ${count ?? 0} بند إلى "سعر قديم"`);
    } catch (e) {
      setCleanupResult(`خطأ: ${(e as Error).message}`);
    } finally {
      setCleanupRunning(false);
    }
  }

  return (
    <div className="p-6 max-w-4xl mx-auto" dir="rtl">
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-slate-900">لوحة الإدارة</h1>
        <p className="text-sm text-slate-500 mt-1">إدارة النظام والحوكمة</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 mb-6">
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-4">
            <ShieldCheck size={18} className="text-emerald-600" />
            <h2 className="text-sm font-semibold text-slate-900">جدران الحوكمة</h2>
            <span className="text-xs text-emerald-600 bg-emerald-100 px-2 py-0.5 rounded-full font-medium mr-auto">
              5 / 5 نشط
            </span>
          </div>
          <div className="space-y-2.5">
            {GOVERNANCE_WALLS.map(wall => (
              <div key={wall.number} className="flex items-start gap-3 p-3 bg-emerald-50 border border-emerald-100 rounded-lg">
                <CheckCircle size={16} className="text-emerald-600 mt-0.5 flex-shrink-0" />
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-emerald-700">الجدار {wall.number}</span>
                    <span className="text-xs font-semibold text-slate-800">{wall.name}</span>
                  </div>
                  <p className="text-xs text-slate-500 mt-0.5">{wall.description}</p>
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-4">
            <Activity size={18} className="text-blue-600" />
            <h2 className="text-sm font-semibold text-slate-900">إجراءات الصيانة</h2>
          </div>

          <div className="space-y-3">
            <div className="p-4 border border-slate-200 rounded-xl">
              <div className="flex items-start gap-3 mb-3">
                <AlertTriangle size={16} className="text-amber-500 mt-0.5 flex-shrink-0" />
                <div>
                  <p className="text-sm font-semibold text-slate-800">تنظيف الأسعار القديمة</p>
                  <p className="text-xs text-slate-500 mt-0.5">
                    تحديث البنود المعتمدة منذ أكثر من 90 يوماً إلى حالة "سعر قديم" لإعادة تقييمها
                  </p>
                </div>
              </div>
              <button
                onClick={handleStaleCleanup}
                disabled={cleanupRunning}
                className="flex items-center gap-2 px-4 py-2 bg-amber-500 hover:bg-amber-600 disabled:bg-amber-300 text-white text-sm font-medium rounded-lg transition-colors"
              >
                <RefreshCw size={14} className={cleanupRunning ? 'animate-spin' : ''} />
                <span>{cleanupRunning ? 'جاري التنظيف...' : 'تشغيل التنظيف'}</span>
              </button>
              {cleanupResult && (
                <p className="text-xs mt-2 text-slate-600">{cleanupResult}</p>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-4">
            <Database size={18} className="text-slate-600" />
            <h2 className="text-sm font-semibold text-slate-900">سجل الترحيلات</h2>
            {migrations.length > 0 && (
              <span className="text-xs text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full mr-auto">
                {migrations.length} ترحيل
              </span>
            )}
          </div>

          {loading ? (
            <div className="space-y-2">
              {Array.from({ length: 4 }).map((_, i) => (
                <div key={i} className="skeleton h-8 rounded" />
              ))}
            </div>
          ) : migrations.length === 0 ? (
            <div className="py-8 text-center">
              <Database size={28} className="mx-auto text-slate-300 mb-2" />
              <p className="text-xs text-slate-500">لا توجد بيانات</p>
            </div>
          ) : (
            <div className="space-y-1.5 max-h-64 overflow-y-auto">
              {migrations.map((m, i) => (
                <div key={i} className="flex items-center gap-3 px-3 py-2 bg-slate-50 rounded-lg">
                  <CheckCircle size={13} className="text-emerald-500 flex-shrink-0" />
                  <div className="flex-1 min-w-0">
                    <p className="text-xs font-medium text-slate-700 truncate">{m.name || m.version}</p>
                  </div>
                  <span className="text-xs text-slate-400 flex-shrink-0 ltr">
                    {m.inserted_at ? new Date(m.inserted_at).toLocaleDateString('en-GB') : ''}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="flex items-center gap-2 mb-4">
            <Clock size={18} className="text-blue-600" />
            <h2 className="text-sm font-semibold text-slate-900">سجل التسعير</h2>
            {auditLog.length > 0 && (
              <span className="text-xs text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full mr-auto">
                آخر 20 سجل
              </span>
            )}
          </div>

          {loading ? (
            <div className="space-y-2">
              {Array.from({ length: 4 }).map((_, i) => (
                <div key={i} className="skeleton h-8 rounded" />
              ))}
            </div>
          ) : auditLog.length === 0 ? (
            <div className="py-8 text-center">
              <Clock size={28} className="mx-auto text-slate-300 mb-2" />
              <p className="text-xs text-slate-500">لا توجد سجلات تسعير بعد</p>
            </div>
          ) : (
            <div className="space-y-1.5 max-h-64 overflow-y-auto">
              {auditLog.map(entry => (
                <div key={entry.id} className="flex items-center gap-3 px-3 py-2 bg-slate-50 rounded-lg">
                  <div className={`w-2 h-2 rounded-full flex-shrink-0 ${
                    entry.override_type === 'manual' ? 'bg-blue-500' :
                    entry.source_type === 'Approved' ? 'bg-emerald-500' : 'bg-amber-400'
                  }`} />
                  <div className="flex-1 min-w-0">
                    <p className="text-xs text-slate-700 truncate ltr">{entry.boq_item_id.slice(0, 8)}...</p>
                  </div>
                  <span className="text-xs font-medium text-slate-700 ltr flex-shrink-0">{sarFormat(entry.unit_rate)}</span>
                  <span className="text-xs text-slate-400 flex-shrink-0 ltr">
                    {new Date(entry.created_at).toLocaleDateString('en-GB')}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
