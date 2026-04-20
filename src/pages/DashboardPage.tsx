import { useState, useEffect } from 'react';
import { FolderOpen, FileSpreadsheet, TrendingUp, DollarSign, Plus, ArrowLeft } from 'lucide-react';
import type { LucideIcon } from 'lucide-react';
import { supabase } from '../lib/supabase';
import type { Project, BOQFile } from '../types';

interface DashboardPageProps {
  onNavigate: (page: string, data?: unknown) => void;
}

const sarFormat = (v: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR', maximumFractionDigits: 0 }).format(v);

interface Stats {
  totalProjects: number;
  totalBOQFiles: number;
  totalItemsPriced: number;
  totalSARValue: number;
}

interface RecentProject extends Project {
  boq_count: number;
  total_items: number;
  priced_items: number;
}

function StatCard({
  icon: Icon,
  label,
  value,
  color,
}: {
  icon: LucideIcon;
  label: string;
  value: string | number;
  color: string;
}) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 flex items-center gap-4">
      <div className={`w-12 h-12 rounded-xl flex items-center justify-center flex-shrink-0 ${color}`}>
        <Icon size={22} className="text-white" />
      </div>
      <div>
        <p className="text-xs text-slate-500">{label}</p>
        <p className="text-xl font-bold text-slate-900 mt-0.5">{value}</p>
      </div>
    </div>
  );
}

export default function DashboardPage({ onNavigate }: DashboardPageProps) {
  const [stats, setStats] = useState<Stats>({ totalProjects: 0, totalBOQFiles: 0, totalItemsPriced: 0, totalSARValue: 0 });
  const [recentProjects, setRecentProjects] = useState<RecentProject[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function load() {
      setLoading(true);
      try {
        const [projectsRes, boqFilesRes, boqItemsRes] = await Promise.all([
          supabase.from('projects').select('*').order('created_at', { ascending: false }),
          supabase.from('boq_files').select('*'),
          supabase.from('boq_items').select('status, unit_rate, quantity, boq_file_id').in('status', ['approved', 'manual']),
        ]);

        const projects = (projectsRes.data ?? []) as Project[];
        const boqFiles = (boqFilesRes.data ?? []) as BOQFile[];
        const pricedItems = boqItemsRes.data ?? [];

        const totalSAR = pricedItems.reduce((sum, i) => sum + ((i.unit_rate ?? 0) * (i.quantity ?? 0)), 0);

        setStats({
          totalProjects: projects.length,
          totalBOQFiles: boqFiles.length,
          totalItemsPriced: pricedItems.length,
          totalSARValue: totalSAR,
        });

        const recent = projects.slice(0, 5).map(p => {
          const pFiles = boqFiles.filter(f => f.project_id === p.id);
          const totalItems = pFiles.reduce((s, f) => s + (f.total_items ?? 0), 0);
          const pricedItemsCount = pFiles.reduce((s, f) => s + (f.priced_items ?? 0), 0);
          return {
            ...p,
            boq_count: pFiles.length,
            total_items: totalItems,
            priced_items: pricedItemsCount,
          };
        });

        setRecentProjects(recent);
      } catch (e) {
        console.error(e);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, []);

  return (
    <div className="p-6 max-w-5xl mx-auto" dir="rtl">
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-slate-900">لوحة التحكم</h1>
        <p className="text-sm text-slate-500 mt-1">نظرة عامة على مشاريع التسعير</p>
      </div>

      {loading ? (
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
          {Array.from({ length: 4 }).map((_, i) => (
            <div key={i} className="bg-white rounded-xl border border-slate-200 p-5 h-24">
              <div className="skeleton h-full rounded-lg" />
            </div>
          ))}
        </div>
      ) : (
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
          <StatCard icon={FolderOpen} label="إجمالي المشاريع" value={stats.totalProjects} color="bg-blue-600" />
          <StatCard icon={FileSpreadsheet} label="ملفات BOQ" value={stats.totalBOQFiles} color="bg-slate-700" />
          <StatCard icon={TrendingUp} label="بنود مُسعَّرة" value={stats.totalItemsPriced.toLocaleString('en-US')} color="bg-emerald-600" />
          <StatCard icon={DollarSign} label="القيمة الإجمالية" value={sarFormat(stats.totalSARValue)} color="bg-amber-500" />
        </div>
      )}

      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-100">
          <h2 className="text-sm font-semibold text-slate-900">أحدث المشاريع</h2>
          <button
            onClick={() => onNavigate('projects')}
            className="flex items-center gap-1.5 text-xs text-blue-600 hover:text-blue-800 font-medium"
          >
            <span>عرض الكل</span>
            <ArrowLeft size={14} />
          </button>
        </div>

        {loading ? (
          <div className="divide-y divide-slate-100">
            {Array.from({ length: 3 }).map((_, i) => (
              <div key={i} className="px-5 py-4 flex items-center gap-4">
                <div className="skeleton h-4 rounded w-32" />
                <div className="flex-1 skeleton h-3 rounded" />
              </div>
            ))}
          </div>
        ) : recentProjects.length === 0 ? (
          <div className="py-12 text-center">
            <FolderOpen size={36} className="mx-auto text-slate-300 mb-3" />
            <p className="text-sm text-slate-500">لا توجد مشاريع بعد</p>
            <button
              onClick={() => onNavigate('projects')}
              className="mt-3 text-sm text-blue-600 hover:text-blue-800 font-medium"
            >
              إنشاء مشروع جديد
            </button>
          </div>
        ) : (
          <div className="divide-y divide-slate-100">
            {recentProjects.map(project => {
              const pct = project.total_items > 0
                ? Math.round((project.priced_items / project.total_items) * 100)
                : 0;
              return (
                <div
                  key={project.id}
                  onClick={() => onNavigate('projects')}
                  className="px-5 py-4 hover:bg-slate-50 cursor-pointer transition-colors"
                >
                  <div className="flex items-center justify-between mb-2">
                    <div>
                      <p className="text-sm font-medium text-slate-800">{project.name}</p>
                      <p className="text-xs text-slate-500 mt-0.5">
                        {project.client} — {project.city}
                      </p>
                    </div>
                    <div className="text-left">
                      <p className="text-xs text-slate-500">{project.boq_count} ملف BOQ</p>
                      <p className="text-xs font-semibold text-emerald-600 mt-0.5">{pct}% مُسعَّر</p>
                    </div>
                  </div>
                  <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-emerald-500 rounded-full transition-all duration-500"
                      style={{ width: `${pct}%` }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div className="mt-6">
        <button
          onClick={() => onNavigate('projects')}
          className="flex items-center gap-2 px-5 py-3 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-xl transition-colors shadow-md shadow-blue-200"
        >
          <Plus size={18} />
          <span>رفع ملف BOQ جديد</span>
        </button>
      </div>
    </div>
  );
}
