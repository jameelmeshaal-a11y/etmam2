import RateLibraryTable from '../components/RateLibraryTable';

interface RateLibraryPageProps {
  onNavigate?: (page: string, data?: unknown) => void;
}

export default function RateLibraryPage({ onNavigate: _onNavigate }: RateLibraryPageProps) {
  return (
    <div className="flex flex-col h-full" dir="rtl">
      <div className="bg-white border-b border-slate-200 px-6 py-4 flex-shrink-0">
        <h1 className="text-xl font-bold text-slate-900">مكتبة الأسعار</h1>
        <p className="text-xs text-slate-500 mt-0.5">إدارة معدلات التسعير القياسية</p>
      </div>
      <div className="flex-1 overflow-hidden">
        <RateLibraryTable />
      </div>
    </div>
  );
}
