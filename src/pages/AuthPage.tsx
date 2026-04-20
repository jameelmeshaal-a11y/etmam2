import { useState } from 'react';
import { HardHat, Mail, Lock, Eye, EyeOff, UserPlus, LogIn } from 'lucide-react';
import { supabase } from '../lib/supabase';
import type { User } from '@supabase/supabase-js';

interface AuthPageProps {
  onAuth: (user: User) => void;
}

type Tab = 'login' | 'register';

export default function AuthPage({ onAuth }: AuthPageProps) {
  const [tab, setTab] = useState<Tab>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!email.trim() || !password.trim()) {
      setError('يرجى إدخال البريد الإلكتروني وكلمة المرور');
      return;
    }

    setLoading(true);
    setError(null);
    setSuccess(null);

    try {
      if (tab === 'login') {
        const { error: authError } = await supabase.auth.signInWithPassword({ email, password });
        if (authError) throw authError;
      } else {
        const { data, error: authError } = await supabase.auth.signUp({ email, password });
        if (authError) throw authError;
        if (data.user && !data.session) {
          setSuccess('تم إرسال رابط التحقق إلى بريدك الإلكتروني. يرجى التحقق منه للمتابعة.');
        }
      }
    } catch (e) {
      const msg = (e as Error).message;
      if (msg.includes('Invalid login')) setError('بريد إلكتروني أو كلمة مرور غير صحيحة');
      else if (msg.includes('Email already registered')) setError('هذا البريد الإلكتروني مسجل مسبقاً');
      else if (msg.includes('Password should be')) setError('كلمة المرور يجب أن تكون 6 أحرف على الأقل');
      else setError(msg);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4" dir="rtl">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="w-16 h-16 rounded-2xl bg-blue-600 flex items-center justify-center mx-auto mb-4 shadow-lg shadow-blue-900/40">
            <HardHat size={32} className="text-white" />
          </div>
          <h1 className="text-2xl font-bold text-white">محرك تسعير المشاريع</h1>
          <p className="text-slate-400 text-sm mt-2">منصة تسعير BOQ للمشاريع الإنشائية</p>
        </div>

        <div className="bg-white rounded-2xl shadow-2xl overflow-hidden">
          <div className="flex">
            <button
              onClick={() => { setTab('login'); setError(null); setSuccess(null); }}
              className={`flex-1 py-3 text-sm font-semibold transition-colors ${
                tab === 'login'
                  ? 'bg-blue-600 text-white'
                  : 'bg-slate-50 text-slate-500 hover:text-slate-700'
              }`}
            >
              <span className="flex items-center justify-center gap-2">
                <LogIn size={15} />
                تسجيل الدخول
              </span>
            </button>
            <button
              onClick={() => { setTab('register'); setError(null); setSuccess(null); }}
              className={`flex-1 py-3 text-sm font-semibold transition-colors ${
                tab === 'register'
                  ? 'bg-blue-600 text-white'
                  : 'bg-slate-50 text-slate-500 hover:text-slate-700'
              }`}
            >
              <span className="flex items-center justify-center gap-2">
                <UserPlus size={15} />
                إنشاء حساب
              </span>
            </button>
          </div>

          <form onSubmit={handleSubmit} className="p-6 space-y-4">
            {error && (
              <div className="p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">
                {error}
              </div>
            )}
            {success && (
              <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-sm text-emerald-700">
                {success}
              </div>
            )}

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1.5">البريد الإلكتروني</label>
              <div className="relative">
                <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">
                  <Mail size={16} />
                </div>
                <input
                  type="email"
                  value={email}
                  onChange={e => setEmail(e.target.value)}
                  placeholder="example@company.com"
                  dir="ltr"
                  className="w-full pr-10 pl-4 py-2.5 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-right"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1.5">كلمة المرور</label>
              <div className="relative">
                <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">
                  <Lock size={16} />
                </div>
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  placeholder="••••••••"
                  dir="ltr"
                  className="w-full pr-10 pl-10 py-2.5 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(v => !v)}
                  className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-semibold rounded-lg transition-colors shadow-md shadow-blue-100 mt-2"
            >
              {loading
                ? (tab === 'login' ? 'جاري الدخول...' : 'جاري الإنشاء...')
                : (tab === 'login' ? 'تسجيل الدخول' : 'إنشاء الحساب')
              }
            </button>
          </form>
        </div>

        <p className="text-center text-xs text-slate-500 mt-6">
          منصة إتمام للمشاريع الحكومية — المملكة العربية السعودية
        </p>
      </div>
    </div>
  );
}
