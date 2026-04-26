import { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { FileText, Download, TrendingUp, CheckCircle, Clock, AlertTriangle, BarChart2 } from 'lucide-react';
import { supabase } from '../lib/supabase';
import type { Project, BOQFile } from '../types';

interface BOQSummaryPageProps {
  onNavigate: (page: string, data?: unknown) => void;
}

interface FileSummary extends BOQFile {
  actual_priced: number;
  actual_total: number;
  pending_count: number;
  needs_review_count: number;
  descriptive_count: number;
  total_value: number;
  is_complete: boolean;
}

interface ProjectSummary extends Project {
  files: FileSummary[];
  total_items: number;
  priced_items: number;
  total_value: number;
  pct: number;
}

const sarFormat = (v: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'SAR', maximumFractionDigits: 0 }).format(v);

export default function BOQSummaryPage({ onNavigate }: BOQSummaryPageProps) {
  const [projects, setProjects] = useState<Project[]>([]);
  const [selectedProjectId, setSelectedProjectId] = useState<string>('');
  const [summary, setSummary] = useState<ProjectSummary | null>(null);
  const [loading, setLoading] = useState(false);
  const [projectsLoading, setProjectsLoading] = useState(true);

  useEffect(() => {
    async function loadProjects() {
      setProjectsLoading(true);
      const { data } = await supabase
        .from('projects')
        .select('*')
        .order('created_at', { ascending: false });
      const list = (data ?? []) as Project[];
      setProjects(list);
      if (list.length > 0) setSelectedProjectId(list[0].id);
      setProjectsLoading(false);
    }
    loadProjects();
  }, []);

  useEffect(() => {
    if (!selectedProjectId) { setSummary(null); return; }
    loadSummary(selectedProjectId);
  }, [selectedProjectId]);

  async function loadSummary(projectId: string) {
    setLoading(true);
    try {
      const project = projects.find(p => p.id === projectId);
      if (!project) return;

      const { data: filesData } = await supabase
        .from('boq_files')
        .select('*')
        .eq('project_id', projectId)
        .eq('is_archived', false)
        .order('created_at', { ascending: true });

      const files = (filesData ?? []) as BOQFile[];

      const { data: itemsData } = await supabase
        .from('boq_items')
        .select('boq_file_id, status, unit_rate, quantity, total_price, override_type')
        .in('boq_file_id', files.map(f => f.id));

      const items = itemsData ?? [];

      const fileSummaries: FileSummary[] = files.map(file => {
        const fileItems = items.filter(i => i.boq_file_id === file.id);
        const priced = fileItems.filter(i =>
          (i.status === 'approved' || i.override_type === 'manual') && i.unit_rate != null
        );
        const pending = fileItems.filter(i => i.status === 'pending');
        const needsReview = fileItems.filter(i => i.status === 'needs_review');
        const descriptive = fileItems.filter(i => i.status === 'descriptive');
        const totalValue = priced.reduce((s, i) => s + (i.total_price ?? (i.unit_rate ?? 0) * (i.quantity ?? 0)), 0);
        const nonDescriptive = fileItems.filter(i => i.status !== 'descriptive');

        return {
          ...file,
          actual_priced: priced.length,
          actual_total: nonDescriptive.length,
          pending_count: pending.length,
          needs_review_count: needsReview.length,
          descriptive_count: descriptive.length,
          total_value: totalValue,
          is_complete: nonDescriptive.length > 0 && priced.length === nonDescriptive.length,
        };
      });

      const totalItems = fileSummaries.reduce((s, f) => s + f.actual_total, 0);
      const pricedItems = fileSummaries.reduce((s, f) => s + f.actual_priced, 0);
      const totalValue = fileSummaries.reduce((s, f) => s + f.total_value, 0);

      setSummary({
        ...project,
        files: fileSummaries,
        total_items: totalItems,
        priced_items: pricedItems,
        total_value: totalValue,
        pct: totalItems > 0 ? Math.round((pricedItems / totalItems) * 100) : 0,
      });
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  }

  async function exportToExcel() {
    if (!summary) return;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('ملخص جداول الكميات', { views: [{ rightToLeft: true }] });

    ws.columns = [
      { width: 40 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 20 },
      { width: 15 },
    ];

    const titleRow = ws.addRow([`ملخص جداول الكميات — ${summary.name}`]);
    titleRow.font = { bold: true, size: 14 };
    titleRow.alignment = { horizontal: 'right' };
    ws.mergeCells(`A${titleRow.number}:F${titleRow.number}`);

    const subRow = ws.addRow([`الجهة: ${summary.client || '—'} | المدينة: ${summary.city || '—'} | تاريخ التقرير: ${new Date().toLocaleDateString('en-GB')}`]);
    subRow.font = { size: 10, color: { argb: 'FF64748B' } };
    subRow.alignment = { horizontal: 'right' };
    ws.mergeCells(`A${subRow.number}:F${subRow.number}`);

    ws.addRow([]);

    const headerRow = ws.addRow(['جدول الكميات', 'عدد البنود', 'المُسعَّرة', 'في الانتظار', 'الإجمالي (ر.س)', 'الحالة']);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { horizontal: 'center' };
    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
    });

    for (const file of summary.files) {
      const row = ws.addRow([
        file.name,
        file.actual_total,
        file.actual_priced,
        file.pending_count + file.needs_review_count,
        file.total_value,
        file.is_complete ? 'مكتمل' : 'جاري',
      ]);
      row.getCell(5).numFmt = '#,##0.00';
      row.getCell(6).font = {
        color: { argb: file.is_complete ? 'FF059669' : 'FFD97706' },
        bold: true,
      };
      row.alignment = { horizontal: 'center' };
      row.getCell(1).alignment = { horizontal: 'right' };
    }

    ws.addRow([]);
    const totalRow = ws.addRow([
      'الإجمالي',
      summary.total_items,
      summary.priced_items,
      summary.total_items - summary.priced_items,
      summary.total_value,
      `${summary.pct}%`,
    ]);
    totalRow.font = { bold: true };
    totalRow.getCell(5).numFmt = '#,##0.00';
    totalRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ملخص_${summary.name}_${new Date().toLocaleDateString('en-GB').replace(/\//g, '-')}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  const selectedProject = projects.find(p => p.id === selectedProjectId);

  return (
    <div className="p-6 max-w-5xl mx-auto" dir="rtl">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center">
            <BarChart2 size={20} className="text-emerald-600" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-900">تقرير ملخص المشروع</h1>
            <p className="text-sm text-slate-500 mt-0.5">ملخص جداول الكميات والتسعير</p>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5 mb-6">
        <label className="block text-xs font-semibold text-slate-700 mb-2">اختر المشروع</label>
        {projectsLoading ? (
          <div className="h-10 skeleton rounded-lg" />
        ) : (
          <select
            value={selectedProjectId}
            onChange={e => setSelectedProjectId(e.target.value)}
            className="w-full px-3 py-2.5 text-sm border-2 border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 bg-white text-slate-800"
          >
            {projects.length === 0 && <option value="">لا توجد مشاريع</option>}
            {projects.map(p => (
              <option key={p.id} value={p.id}>{p.name}</option>
            ))}
          </select>
        )}
      </div>

      {loading ? (
        <div className="space-y-3">
          {Array.from({ length: 4 }).map((_, i) => (
            <div key={i} className="bg-white rounded-xl border border-slate-200 p-5 h-16 skeleton" />
          ))}
        </div>
      ) : summary ? (
        <>
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
            <SummaryCard
              icon={<FileText size={20} className="text-blue-600" />}
              bg="bg-blue-50"
              label="إجمالي البنود"
              value={summary.total_items.toLocaleString('en-US')}
            />
            <SummaryCard
              icon={<CheckCircle size={20} className="text-emerald-600" />}
              bg="bg-emerald-50"
              label="البنود المُسعَّرة"
              value={summary.priced_items.toLocaleString('en-US')}
            />
            <SummaryCard
              icon={<TrendingUp size={20} className="text-amber-600" />}
              bg="bg-amber-50"
              label="نسبة التسعير"
              value={`${summary.pct}%`}
            />
            <SummaryCard
              icon={<BarChart2 size={20} className="text-slate-600" />}
              bg="bg-slate-100"
              label="القيمة الإجمالية"
              value={sarFormat(summary.total_value)}
              small
            />
          </div>

          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <FileText size={16} className="text-slate-500" />
                <span className="text-sm font-semibold text-slate-800">
                  ملخص جداول الكميات — {selectedProject?.name}
                </span>
              </div>
              <button
                onClick={exportToExcel}
                className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-medium rounded-lg transition-colors"
              >
                <Download size={14} />
                <span>Excel</span>
              </button>
            </div>

            {summary.files.length === 0 ? (
              <div className="py-16 text-center">
                <FileText size={36} className="mx-auto text-slate-300 mb-3" />
                <p className="text-sm text-slate-500">لا توجد ملفات BOQ لهذا المشروع</p>
                <button
                  onClick={() => onNavigate('projects')}
                  className="mt-3 text-sm text-blue-600 hover:text-blue-800 font-medium"
                >
                  رفع ملف BOQ
                </button>
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm" dir="rtl">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-100">
                      <th className="px-5 py-3 text-xs font-semibold text-slate-600 text-right">جدول الكميات</th>
                      <th className="px-4 py-3 text-xs font-semibold text-slate-600 text-center">عدد البنود</th>
                      <th className="px-4 py-3 text-xs font-semibold text-slate-600 text-center">المُسعَّرة</th>
                      <th className="px-4 py-3 text-xs font-semibold text-slate-600 text-center">في الانتظار</th>
                      <th className="px-4 py-3 text-xs font-semibold text-slate-600 text-left">الإجمالي</th>
                      <th className="px-4 py-3 text-xs font-semibold text-slate-600 text-center">الحالة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {summary.files.map(file => {
                      const filePct = file.actual_total > 0
                        ? Math.round((file.actual_priced / file.actual_total) * 100)
                        : 0;
                      const pending = file.pending_count + file.needs_review_count;
                      return (
                        <tr
                          key={file.id}
                          onClick={() => onNavigate('boq', { boqFileId: file.id, boqFile: file })}
                          className="hover:bg-slate-50 cursor-pointer transition-colors"
                        >
                          <td className="px-5 py-4">
                            <p className="text-sm font-medium text-slate-800 truncate max-w-xs">{file.name}</p>
                            <div className="flex items-center gap-2 mt-1.5">
                              <div className="flex-1 h-1.5 bg-slate-100 rounded-full overflow-hidden max-w-32">
                                <div
                                  className="h-full bg-emerald-500 rounded-full transition-all"
                                  style={{ width: `${filePct}%` }}
                                />
                              </div>
                              <span className="text-xs text-slate-400">{filePct}%</span>
                            </div>
                          </td>
                          <td className="px-4 py-4 text-center">
                            <span className="text-sm font-medium text-slate-700">{file.actual_total.toLocaleString('en-US')}</span>
                          </td>
                          <td className="px-4 py-4 text-center">
                            <span className="inline-flex items-center gap-1 text-sm font-semibold text-emerald-700">
                              <CheckCircle size={13} />
                              {file.actual_priced.toLocaleString('en-US')}
                            </span>
                          </td>
                          <td className="px-4 py-4 text-center">
                            {pending > 0 ? (
                              <span className="inline-flex items-center gap-1 text-sm text-amber-600">
                                <Clock size={13} />
                                {pending.toLocaleString('en-US')}
                              </span>
                            ) : (
                              <span className="text-sm text-slate-300">—</span>
                            )}
                          </td>
                          <td className="px-4 py-4 text-left">
                            <span className="text-sm font-semibold text-slate-900 ltr">
                              {sarFormat(file.total_value)}
                            </span>
                          </td>
                          <td className="px-4 py-4 text-center">
                            {file.is_complete ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-50 text-emerald-700 text-xs font-semibold rounded-full border border-emerald-200">
                                <CheckCircle size={11} />
                                مكتمل
                              </span>
                            ) : file.needs_review_count > 0 ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-red-50 text-red-600 text-xs font-semibold rounded-full border border-red-200">
                                <AlertTriangle size={11} />
                                يحتاج مراجعة
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-amber-50 text-amber-700 text-xs font-semibold rounded-full border border-amber-200">
                                <Clock size={11} />
                                جاري
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr className="border-t-2 border-slate-200 bg-slate-50">
                      <td className="px-5 py-4">
                        <span className="text-sm font-bold text-slate-900">الإجمالي</span>
                      </td>
                      <td className="px-4 py-4 text-center">
                        <span className="text-sm font-bold text-slate-900">{summary.total_items.toLocaleString('en-US')}</span>
                      </td>
                      <td className="px-4 py-4 text-center">
                        <span className="text-sm font-bold text-emerald-700">{summary.priced_items.toLocaleString('en-US')}</span>
                      </td>
                      <td className="px-4 py-4 text-center">
                        <span className="text-sm font-bold text-amber-600">
                          {(summary.total_items - summary.priced_items).toLocaleString('en-US')}
                        </span>
                      </td>
                      <td className="px-4 py-4 text-left">
                        <span className="text-sm font-bold text-slate-900 ltr">{sarFormat(summary.total_value)}</span>
                      </td>
                      <td className="px-4 py-4 text-center">
                        <span className="text-sm font-bold text-slate-700">{summary.pct}%</span>
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            )}
          </div>
        </>
      ) : null}
    </div>
  );
}

function SummaryCard({
  icon,
  bg,
  label,
  value,
  small,
}: {
  icon: React.ReactNode;
  bg: string;
  label: string;
  value: string;
  small?: boolean;
}) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 flex items-center gap-4">
      <div className={`w-11 h-11 rounded-xl flex items-center justify-center flex-shrink-0 ${bg}`}>
        {icon}
      </div>
      <div className="min-w-0">
        <p className="text-xs text-slate-500">{label}</p>
        <p className={`font-bold text-slate-900 mt-0.5 truncate ${small ? 'text-base' : 'text-xl'}`}>{value}</p>
      </div>
    </div>
  );
}
