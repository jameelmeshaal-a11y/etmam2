import { useState } from 'react';
import { LayoutDashboard, FolderOpen, BookOpen, Settings, LogOut, HardHat, BarChart2, ChevronRight, ChevronLeft } from 'lucide-react';
import type { LucideIcon } from 'lucide-react';
import type { User } from '@supabase/supabase-js';
import { supabase } from '../lib/supabase';

type Page = 'dashboard' | 'projects' | 'boq' | 'library' | 'admin' | 'summary';

interface NavItem {
  id: Page;
  label: string;
  icon: LucideIcon;
}

const NAV_ITEMS: NavItem[] = [
  { id: 'dashboard', label: 'لوحة التحكم', icon: LayoutDashboard },
  { id: 'projects', label: 'المشاريع', icon: FolderOpen },
  { id: 'summary', label: 'ملخص التقارير', icon: BarChart2 },
  { id: 'library', label: 'مكتبة الأسعار', icon: BookOpen },
  { id: 'admin', label: 'لوحة الإدارة', icon: Settings },
];

interface LayoutProps {
  currentPage: string;
  onNavigate: (page: Page) => void;
  user: User;
  children: React.ReactNode;
}

export default function Layout({ currentPage, onNavigate, user, children }: LayoutProps) {
  const activePage = currentPage as Page;
  const [collapsed, setCollapsed] = useState(false);

  async function handleLogout() {
    await supabase.auth.signOut();
  }

  return (
    <div className="flex h-screen bg-slate-50 overflow-hidden" dir="rtl">
      <aside
        className={`relative flex-shrink-0 bg-slate-900 flex flex-col shadow-xl transition-all duration-300 ease-in-out ${
          collapsed ? 'w-[60px]' : 'w-64'
        }`}
      >
        <div className={`px-3 py-5 border-b border-slate-700/50 flex items-center ${collapsed ? 'justify-center' : 'gap-3'}`}>
          <div className="w-9 h-9 rounded-lg bg-blue-500 flex items-center justify-center flex-shrink-0">
            <HardHat size={20} className="text-white" />
          </div>
          {!collapsed && (
            <div className="overflow-hidden">
              <h1 className="text-white font-bold text-sm leading-tight whitespace-nowrap">محرك تسعير المشاريع</h1>
              <p className="text-slate-400 text-xs mt-0.5">منصة إتمام</p>
            </div>
          )}
        </div>

        <nav className="flex-1 px-2 py-4 space-y-1 overflow-y-auto overflow-x-hidden">
          {NAV_ITEMS.map(item => {
            const Icon = item.icon;
            const isActive = activePage === item.id;
            return (
              <button
                key={item.id}
                onClick={() => onNavigate(item.id)}
                title={collapsed ? item.label : undefined}
                className={`w-full flex items-center gap-3 px-2.5 py-2.5 rounded-lg text-sm font-medium transition-all duration-150 ${
                  collapsed ? 'justify-center' : 'text-right'
                } ${
                  isActive
                    ? 'bg-blue-600 text-white shadow-md shadow-blue-900/30'
                    : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                }`}
              >
                <Icon size={18} className={`flex-shrink-0 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                {!collapsed && <span className="whitespace-nowrap">{item.label}</span>}
              </button>
            );
          })}
        </nav>

        <div className="px-2 py-4 border-t border-slate-700/50">
          {!collapsed && (
            <div className="flex items-center gap-3 mb-3 px-1">
              <div className="w-8 h-8 rounded-full bg-slate-700 flex items-center justify-center flex-shrink-0">
                <span className="text-white text-xs font-semibold">
                  {user.email?.charAt(0).toUpperCase() ?? 'U'}
                </span>
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-slate-200 text-xs font-medium truncate">{user.email}</p>
                <p className="text-slate-500 text-xs">مستخدم نظام</p>
              </div>
            </div>
          )}
          <button
            onClick={handleLogout}
            title={collapsed ? 'تسجيل الخروج' : undefined}
            className={`w-full flex items-center gap-2 px-2.5 py-2 rounded-lg text-sm text-slate-400 hover:bg-slate-800 hover:text-red-400 transition-all ${
              collapsed ? 'justify-center' : ''
            }`}
          >
            <LogOut size={16} className="flex-shrink-0" />
            {!collapsed && <span>تسجيل الخروج</span>}
          </button>
        </div>

        <button
          onClick={() => setCollapsed(prev => !prev)}
          className="absolute -left-3 top-1/2 -translate-y-1/2 w-6 h-6 rounded-full bg-slate-700 border border-slate-600 flex items-center justify-center text-slate-300 hover:bg-blue-600 hover:text-white hover:border-blue-500 transition-all duration-150 shadow-md z-10"
          title={collapsed ? 'توسيع القائمة' : 'طي القائمة'}
        >
          {collapsed ? <ChevronLeft size={13} /> : <ChevronRight size={13} />}
        </button>
      </aside>

      <main className="flex-1 overflow-y-auto bg-slate-50 min-w-0">
        {children}
      </main>
    </div>
  );
}
