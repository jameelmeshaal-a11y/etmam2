import { useState, useEffect, useRef, useCallback } from 'react';
import { Plus, Trash2, FolderOpen, FileSpreadsheet, Upload, ChevronDown, ChevronUp, AlertTriangle, CheckCircle, Loader, Archive, ArchiveRestore } from 'lucide-react';
import { supabase } from '../lib/supabase';
import { parseEtimadBOQ } from '../lib/excelParser';
import type { Project, BOQFile } from '../types';

interface ProjectsPageProps {
  onNavigate: (page: string, data?: unknown) => void;
}

interface UploadState {
  phase: 'idle' | 'parsing' | 'uploading' | 'saving' | 'pricing' | 'done' | 'error';
  message: string;
}

interface ProjectWithFiles extends Project {
  boq_files: BOQFile[];
  expanded: boolean;
  showArchived: boolean;
}

export default function ProjectsPage({ onNavigate }: ProjectsPageProps) {
  const [projects, setProjects] = useState<ProjectWithFiles[]>([]);
  const [loading, setLoading] = useState(true);
  const [showCreateForm, setShowCreateForm] = useState(false);
  const [newProjectName, setNewProjectName] = useState('');
  const [newProjectClient, setNewProjectClient] = useState('');
  const [newProjectCity, setNewProjectCity] = useState('');
  const [creating, setCreating] = useState(false);
  const [createError, setCreateError] = useState<string | null>(null);
  const [uploadStates, setUploadStates] = useState<Record<string, UploadState>>({});
  const fileInputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const loadProjects = useCallback(async () => {
    setLoading(true);
    try {
      const { data: pData } = await supabase
        .from('projects')
        .select('*')
        .order('created_at', { ascending: false });
      const projectList = (pData ?? []) as Project[];

      const { data: fData } = await supabase
        .from('boq_files')
        .select('*')
        .order('created_at', { ascending: false });
      const files = (fData ?? []) as BOQFile[];

      const enriched: ProjectWithFiles[] = projectList.map(p => ({
        ...p,
        boq_files: files.filter(f => f.project_id === p.id),
        expanded: false,
        showArchived: false,
      }));

      setProjects(enriched);
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { loadProjects(); }, [loadProjects]);

  function setUploadState(projectId: string, state: UploadState) {
    setUploadStates(prev => ({ ...prev, [projectId]: state }));
  }

  async function handleCreateProject(e: React.FormEvent) {
    e.preventDefault();
    if (!newProjectName.trim()) { setCreateError('اسم المشروع مطلوب'); return; }
    setCreating(true);
    setCreateError(null);
    try {
      const { data: userData } = await supabase.auth.getUser();
      const { data, error } = await supabase
        .from('projects')
        .insert({
          name: newProjectName.trim(),
          client: newProjectClient.trim(),
          city: newProjectCity.trim(),
          created_by: userData.user?.id,
        })
        .select()
        .single();
      if (error) throw error;
      setNewProjectName('');
      setNewProjectClient('');
      setNewProjectCity('');
      setShowCreateForm(false);
      setProjects(prev => [{
        ...(data as Project),
        boq_files: [],
        expanded: true,
        showArchived: false,
      }, ...prev]);
    } catch (e) {
      setCreateError((e as Error).message);
    } finally {
      setCreating(false);
    }
  }

  async function handleDeleteProject(projectId: string) {
    if (!confirm('هل أنت متأكد من حذف هذا المشروع؟ سيتم حذف جميع ملفات BOQ المرتبطة به.')) return;
    try {
      await supabase.from('projects').delete().eq('id', projectId);
      setProjects(prev => prev.filter(p => p.id !== projectId));
    } catch (e) {
      console.error(e);
    }
  }

  async function handleDeleteBOQFile(projectId: string, fileId: string, e: React.MouseEvent) {
    e.stopPropagation();
    if (!confirm('هل أنت متأكد من حذف هذا الملف؟ سيتم حذف جميع البنود المرتبطة به.')) return;
    try {
      const fileRecord = projects
        .flatMap(p => p.boq_files)
        .find(f => f.id === fileId);

      await supabase.from('boq_items').delete().eq('boq_file_id', fileId);
      await supabase.from('boq_files').delete().eq('id', fileId);

      if (fileRecord?.storage_path) {
        await supabase.storage.from('boq-files').remove([fileRecord.storage_path]);
      }

      setProjects(prev => prev.map(p =>
        p.id === projectId
          ? { ...p, boq_files: p.boq_files.filter(f => f.id !== fileId) }
          : p
      ));
    } catch (e) {
      console.error(e);
    }
  }

  async function handleArchiveBOQFile(projectId: string, fileId: string, currentlyArchived: boolean, e: React.MouseEvent) {
    e.stopPropagation();
    try {
      await supabase
        .from('boq_files')
        .update({
          is_archived: !currentlyArchived,
          archived_at: !currentlyArchived ? new Date().toISOString() : null,
        })
        .eq('id', fileId);
      setProjects(prev => prev.map(p =>
        p.id === projectId
          ? {
              ...p,
              boq_files: p.boq_files.map(f =>
                f.id === fileId
                  ? { ...f, is_archived: !currentlyArchived, archived_at: !currentlyArchived ? new Date().toISOString() : null }
                  : f
              ),
            }
          : p
      ));
    } catch (e) {
      console.error(e);
    }
  }

  function toggleExpand(projectId: string) {
    setProjects(prev => prev.map(p =>
      p.id === projectId ? { ...p, expanded: !p.expanded } : p
    ));
  }

  function toggleShowArchived(projectId: string) {
    setProjects(prev => prev.map(p =>
      p.id === projectId ? { ...p, showArchived: !p.showArchived } : p
    ));
  }

  async function handleFileUpload(projectId: string, file: File) {
    if (!file.name.match(/\.xlsx?$/i)) {
      setUploadState(projectId, { phase: 'error', message: 'يُقبل ملفات Excel فقط (.xlsx)' });
      return;
    }

    setUploadState(projectId, { phase: 'parsing', message: 'جاري تحليل ملف Excel...' });

    let parsedItems: Awaited<ReturnType<typeof parseEtimadBOQ>>['items'];

    try {
      const buffer = await file.arrayBuffer();
      const result = await parseEtimadBOQ(buffer);
      parsedItems = result.items;
      if (parsedItems.length === 0) {
        setUploadState(projectId, { phase: 'error', message: 'لم يتم العثور على بنود في الملف' });
        return;
      }
    } catch (e) {
      setUploadState(projectId, { phase: 'error', message: `خطأ في تحليل الملف: ${(e as Error).message}` });
      return;
    }

    setUploadState(projectId, { phase: 'uploading', message: 'جاري رفع الملف...' });

    let storagePath: string;
    try {
      const { data: userData } = await supabase.auth.getUser();
      const userId = userData.user?.id ?? 'anon';
      const ext = file.name.split('.').pop() ?? 'xlsx';
      storagePath = `${userId}/${projectId}/${Date.now()}.${ext}`;
      const { error: storageError } = await supabase.storage
        .from('boq-files')
        .upload(storagePath, file, { upsert: false });
      if (storageError) throw storageError;
    } catch (e) {
      setUploadState(projectId, { phase: 'error', message: `خطأ في الرفع: ${(e as Error).message}` });
      return;
    }

    setUploadState(projectId, { phase: 'saving', message: 'جاري حفظ البيانات...' });

    let boqFileId: string;
    try {
      const { data: userData } = await supabase.auth.getUser();
      const { data: fileRecord, error: fileError } = await supabase
        .from('boq_files')
        .insert({
          project_id: projectId,
          name: file.name,
          storage_path: storagePath,
          city: projects.find(p => p.id === projectId)?.city ?? '',
          total_items: parsedItems.length,
          priced_items: 0,
          created_by: userData.user?.id,
        })
        .select()
        .single();
      if (fileError) throw fileError;
      boqFileId = (fileRecord as BOQFile).id;

      const itemsToInsert = parsedItems.map(item => ({
        boq_file_id: boqFileId,
        item_no: item.item_no,
        description: item.description,
        unit: item.unit,
        quantity: item.quantity ?? 0,
        row_index: item.row_index,
        status: item.is_descriptive ? 'descriptive' : 'pending',
        override_type: null,
        unit_rate: null,
        total_price: null,
        confidence: 0,
        materials: 0,
        labor: 0,
        equipment: 0,
        logistics: 0,
        risk: 0,
        profit: 0,
      }));

      const chunkSize = 100;
      for (let i = 0; i < itemsToInsert.length; i += chunkSize) {
        const chunk = itemsToInsert.slice(i, i + chunkSize);
        const { error: insertError } = await supabase.from('boq_items').insert(chunk);
        if (insertError) throw insertError;
      }
    } catch (e) {
      setUploadState(projectId, { phase: 'error', message: `خطأ في حفظ البيانات: ${(e as Error).message}` });
      return;
    }

    setUploadState(projectId, { phase: 'done', message: `تم رفع ${parsedItems.length} بند بنجاح. اضغط "تسعير جماعي" للبدء.` });

    await loadProjects();

    const updatedFiles = (await supabase.from('boq_files').select('*').eq('id', boqFileId).single()).data;
    if (updatedFiles) {
      onNavigate('boq', { boqFileId, boqFile: updatedFiles });
    }
  }

  function renderUploadState(projectId: string) {
    const state = uploadStates[projectId];
    if (!state || state.phase === 'idle') return null;

    const isError = state.phase === 'error';
    const isDone = state.phase === 'done';
    const isLoading = !isError && !isDone;

    return (
      <div className={`flex items-center gap-2 p-3 rounded-lg text-xs ${
        isError ? 'bg-red-50 border border-red-200 text-red-700' :
        isDone ? 'bg-emerald-50 border border-emerald-200 text-emerald-700' :
        'bg-blue-50 border border-blue-200 text-blue-700'
      }`}>
        {isLoading && <Loader size={14} className="animate-spin flex-shrink-0" />}
        {isDone && <CheckCircle size={14} className="flex-shrink-0" />}
        {isError && <AlertTriangle size={14} className="flex-shrink-0" />}
        <span>{state.message}</span>
      </div>
    );
  }

  return (
    <div className="p-6 max-w-4xl mx-auto" dir="rtl">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">المشاريع</h1>
          <p className="text-sm text-slate-500 mt-1">{projects.length} مشروع مسجل</p>
        </div>
        <button
          onClick={() => setShowCreateForm(v => !v)}
          className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-xl transition-colors shadow-md shadow-blue-200"
        >
          <Plus size={16} />
          <span>مشروع جديد</span>
        </button>
      </div>

      {showCreateForm && (
        <div className="bg-white rounded-xl border border-blue-200 p-5 mb-5 fade-in">
          <h3 className="text-sm font-semibold text-slate-900 mb-4">إنشاء مشروع جديد</h3>
          <form onSubmit={handleCreateProject} className="space-y-3">
            {createError && (
              <div className="p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">{createError}</div>
            )}
            <div className="grid grid-cols-3 gap-3">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">اسم المشروع *</label>
                <input value={newProjectName} onChange={e => setNewProjectName(e.target.value)} placeholder="مشروع مبنى..." className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">الجهة / العميل</label>
                <input value={newProjectClient} onChange={e => setNewProjectClient(e.target.value)} placeholder="وزارة..." className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">المدينة</label>
                <input value={newProjectCity} onChange={e => setNewProjectCity(e.target.value)} placeholder="الرياض" className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>
            </div>
            <div className="flex items-center gap-3">
              <button type="submit" disabled={creating} className="px-5 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-medium rounded-lg transition-colors">
                {creating ? 'جاري الإنشاء...' : 'إنشاء المشروع'}
              </button>
              <button type="button" onClick={() => setShowCreateForm(false)} className="px-4 py-2 text-sm text-slate-600 hover:bg-slate-100 rounded-lg transition-colors">إلغاء</button>
            </div>
          </form>
        </div>
      )}

      {loading ? (
        <div className="space-y-3">
          {Array.from({ length: 3 }).map((_, i) => (
            <div key={i} className="bg-white rounded-xl border border-slate-200 p-5 h-20">
              <div className="skeleton h-full rounded-lg" />
            </div>
          ))}
        </div>
      ) : projects.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 py-16 text-center">
          <FolderOpen size={48} className="mx-auto text-slate-300 mb-4" />
          <p className="text-sm font-medium text-slate-700">لا توجد مشاريع</p>
          <p className="text-xs text-slate-500 mt-1">أنشئ مشروعاً جديداً لبدء التسعير</p>
        </div>
      ) : (
        <div className="space-y-3">
          {projects.map(project => {
            const activeFiles = project.boq_files.filter(f => !f.is_archived);
            const archivedFiles = project.boq_files.filter(f => f.is_archived);
            const totalItems = activeFiles.reduce((s, f) => s + (f.total_items ?? 0), 0);
            const pricedItems = activeFiles.reduce((s, f) => s + (f.priced_items ?? 0), 0);
            const pct = totalItems > 0 ? Math.round((pricedItems / totalItems) * 100) : 0;
            const uploadState = uploadStates[project.id];
            const isUploading = uploadState && !['idle', 'done', 'error'].includes(uploadState.phase);

            return (
              <div key={project.id} className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div
                  onClick={() => toggleExpand(project.id)}
                  className="px-5 py-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors"
                >
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-xl bg-blue-50 flex items-center justify-center flex-shrink-0">
                      <FolderOpen size={18} className="text-blue-600" />
                    </div>
                    <div>
                      <p className="text-sm font-semibold text-slate-900">{project.name}</p>
                      <p className="text-xs text-slate-500 mt-0.5">{project.client || '—'} · {project.city || '—'}</p>
                    </div>
                  </div>

                  <div className="flex items-center gap-4">
                    <div className="text-left">
                      <p className="text-xs text-slate-500">{activeFiles.length} ملف BOQ</p>
                      <p className="text-xs font-medium text-emerald-600">{pct}% مُسعَّر</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={e => { e.stopPropagation(); handleDeleteProject(project.id); }}
                        className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-300 hover:bg-red-50 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={15} />
                      </button>
                      {project.expanded ? <ChevronUp size={16} className="text-slate-400" /> : <ChevronDown size={16} className="text-slate-400" />}
                    </div>
                  </div>
                </div>

                {project.expanded && (
                  <div className="border-t border-slate-100 px-5 py-4 fade-in">
                    <div className="flex items-center justify-between mb-4">
                      <h4 className="text-xs font-semibold text-slate-700">ملفات BOQ</h4>
                      <div>
                        <input
                          ref={el => { fileInputRefs.current[project.id] = el; }}
                          type="file"
                          accept=".xlsx,.xls"
                          className="hidden"
                          onChange={e => {
                            const file = e.target.files?.[0];
                            if (file) handleFileUpload(project.id, file);
                            e.target.value = '';
                          }}
                        />
                        <button
                          onClick={() => fileInputRefs.current[project.id]?.click()}
                          disabled={isUploading}
                          className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-300 text-white text-xs font-medium rounded-lg transition-colors"
                        >
                          <Upload size={13} />
                          <span>رفع ملف BOQ</span>
                        </button>
                      </div>
                    </div>

                    {renderUploadState(project.id)}

                    {activeFiles.length === 0 ? (
                      <div className="py-8 text-center border-2 border-dashed border-slate-200 rounded-xl mt-3">
                        <FileSpreadsheet size={28} className="mx-auto text-slate-300 mb-2" />
                        <p className="text-xs text-slate-500">لا توجد ملفات BOQ. ارفع ملفاً للبدء.</p>
                      </div>
                    ) : (
                      <div className="space-y-2 mt-3">
                        {activeFiles.map(file => {
                          const filePct = (file.total_items ?? 0) > 0
                            ? Math.round(((file.priced_items ?? 0) / file.total_items) * 100)
                            : 0;
                          return (
                            <BOQFileRow
                              key={file.id}
                              file={file}
                              filePct={filePct}
                              onOpen={() => onNavigate('boq', { boqFileId: file.id, boqFile: file })}
                              onDelete={e => handleDeleteBOQFile(project.id, file.id, e)}
                              onArchive={e => handleArchiveBOQFile(project.id, file.id, file.is_archived, e)}
                            />
                          );
                        })}
                      </div>
                    )}

                    {archivedFiles.length > 0 && (
                      <div className="mt-4">
                        <button
                          onClick={() => toggleShowArchived(project.id)}
                          className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-700 transition-colors"
                        >
                          <Archive size={13} />
                          <span>{project.showArchived ? 'إخفاء' : 'عرض'} الأرشيف ({archivedFiles.length})</span>
                          {project.showArchived ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
                        </button>
                        {project.showArchived && (
                          <div className="space-y-2 mt-2">
                            {archivedFiles.map(file => {
                              const filePct = (file.total_items ?? 0) > 0
                                ? Math.round(((file.priced_items ?? 0) / file.total_items) * 100)
                                : 0;
                              return (
                                <BOQFileRow
                                  key={file.id}
                                  file={file}
                                  filePct={filePct}
                                  archived
                                  onOpen={() => onNavigate('boq', { boqFileId: file.id, boqFile: file })}
                                  onDelete={e => handleDeleteBOQFile(project.id, file.id, e)}
                                  onArchive={e => handleArchiveBOQFile(project.id, file.id, file.is_archived, e)}
                                />
                              );
                            })}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

interface BOQFileRowProps {
  file: BOQFile;
  filePct: number;
  archived?: boolean;
  onOpen: () => void;
  onDelete: (e: React.MouseEvent) => void;
  onArchive: (e: React.MouseEvent) => void;
}

function BOQFileRow({ file, filePct, archived, onOpen, onDelete, onArchive }: BOQFileRowProps) {
  return (
    <div
      onClick={onOpen}
      className={`flex items-center gap-3 p-3 rounded-lg cursor-pointer border transition-all ${
        archived
          ? 'bg-slate-50 border-slate-100 opacity-60 hover:opacity-100'
          : 'bg-slate-50 hover:bg-blue-50 border-slate-100 hover:border-blue-200'
      }`}
    >
      <FileSpreadsheet size={16} className={archived ? 'text-slate-400 flex-shrink-0' : 'text-emerald-600 flex-shrink-0'} />
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <p className="text-xs font-medium text-slate-800 truncate">{file.name}</p>
          {archived && (
            <span className="text-xs px-1.5 py-0.5 bg-slate-200 text-slate-500 rounded font-medium flex-shrink-0">مؤرشف</span>
          )}
        </div>
        <div className="flex items-center gap-3 mt-1">
          <div className="flex-1 h-1 bg-slate-200 rounded-full overflow-hidden">
            <div className="h-full bg-emerald-500 rounded-full" style={{ width: `${filePct}%` }} />
          </div>
          <span className="text-xs text-slate-500 flex-shrink-0">{filePct}%</span>
          <span className="text-xs text-slate-400 flex-shrink-0">{file.total_items ?? 0} بند</span>
        </div>
      </div>
      <div className="flex items-center gap-1 flex-shrink-0" onClick={e => e.stopPropagation()}>
        <button
          onClick={onArchive}
          title={archived ? 'إلغاء الأرشفة' : 'أرشفة'}
          className="w-7 h-7 rounded-md flex items-center justify-center text-slate-400 hover:bg-amber-50 hover:text-amber-600 transition-colors"
        >
          {archived ? <ArchiveRestore size={14} /> : <Archive size={14} />}
        </button>
        <button
          onClick={onDelete}
          title="حذف"
          className="w-7 h-7 rounded-md flex items-center justify-center text-slate-400 hover:bg-red-50 hover:text-red-500 transition-colors"
        >
          <Trash2 size={14} />
        </button>
        <span className="text-xs text-blue-600 font-medium px-2">فتح</span>
      </div>
    </div>
  );
}
