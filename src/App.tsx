import { useState, useEffect } from 'react';
import type { User } from '@supabase/supabase-js';
import { supabase } from './lib/supabase';
import Layout from './components/Layout';
import AuthPage from './pages/AuthPage';
import DashboardPage from './pages/DashboardPage';
import ProjectsPage from './pages/ProjectsPage';
import RateLibraryPage from './pages/RateLibraryPage';
import AdminPage from './pages/AdminPage';
import BOQSummaryPage from './pages/BOQSummaryPage';
import BOQTable from './components/BOQTable';
import type { BOQFile } from './types';

type Page = 'dashboard' | 'projects' | 'boq' | 'library' | 'admin' | 'summary';

interface BOQNavData {
  boqFileId: string;
  boqFile: BOQFile;
}

interface NavData {
  boqFileId?: string;
  boqFile?: BOQFile;
}

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [currentPage, setCurrentPage] = useState<Page>('dashboard');
  const [boqNavData, setBOQNavData] = useState<BOQNavData | null>(null);

  useEffect(() => {
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
      setAuthLoading(false);
    });

    return () => subscription.unsubscribe();
  }, []);

  function handleNavigate(page: string, data?: unknown) {
    const typedPage = page as Page;
    if (typedPage === 'boq' && data) {
      const navData = data as NavData;
      if (navData.boqFileId && navData.boqFile) {
        setBOQNavData({ boqFileId: navData.boqFileId, boqFile: navData.boqFile });
      }
    }
    setCurrentPage(typedPage);
  }

  function handleBackFromBOQ() {
    setBOQNavData(null);
    setCurrentPage('projects');
  }

  if (authLoading) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center" dir="rtl">
        <div className="text-center">
          <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
          <p className="text-slate-400 text-sm">جاري التحميل...</p>
        </div>
      </div>
    );
  }

  if (!user) {
    return <AuthPage onAuth={setUser} />;
  }

  function renderPage() {
    if (currentPage === 'boq' && boqNavData) {
      return (
        <BOQTable
          boqFileId={boqNavData.boqFileId}
          boqFile={boqNavData.boqFile}
          onBack={handleBackFromBOQ}
        />
      );
    }

    switch (currentPage) {
      case 'dashboard':
        return <DashboardPage onNavigate={handleNavigate} />;
      case 'projects':
        return <ProjectsPage onNavigate={handleNavigate} />;
      case 'library':
        return <RateLibraryPage onNavigate={handleNavigate} />;
      case 'admin':
        return <AdminPage onNavigate={handleNavigate} />;
      case 'summary':
        return <BOQSummaryPage onNavigate={handleNavigate} />;
      default:
        return <DashboardPage onNavigate={handleNavigate} />;
    }
  }

  const layoutPage = currentPage === 'boq' ? 'projects' : currentPage as Page;

  return (
    <Layout
      currentPage={layoutPage as Page}
      onNavigate={(page: Page) => handleNavigate(page)}
      user={user}
    >
      {renderPage()}
    </Layout>
  );
}
