"use client";

import { useState } from 'react';
import { useAuth, AuthProvider } from '@/context/AuthContext';

function LoginPage() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const { login } = useAuth();

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setSubmitting(true);
    
    const res = await login(email, password);
    if (!res.success) {
      setError(res.message);
      setSubmitting(false);
    }
  };

  return (
    <main className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4 relative overflow-hidden">
      {/* Decorative colored glow overlays */}
      <div className="absolute top-1/4 left-1/4 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[600px] bg-violet-500/5 rounded-full blur-[120px] pointer-events-none"></div>
      <div className="absolute bottom-1/4 right-1/4 translate-x-1/2 translate-y-1/2 w-[600px] h-[600px] bg-rose-500/5 rounded-full blur-[120px] pointer-events-none"></div>

      <div className="w-full max-w-md p-8 space-y-6 bg-white border border-violet-100 rounded-3xl shadow-xl relative overflow-hidden">
        {/* Subtle top brand color accent line */}
        <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-violet-600 to-rose-500"></div>

        <div className="text-center space-y-2">
          {/* Logo or brand icon placeholder */}
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-gradient-to-br from-violet-600 to-rose-500 text-white font-black text-2xl shadow-md mb-2">
            🍇
          </div>
          <h1 className="text-3xl font-black tracking-tight text-slate-900">
            Berries <span className="bg-gradient-to-r from-violet-600 to-rose-500 bg-clip-text text-transparent">Paradise</span>
          </h1>
          <p className="text-slate-500 text-sm font-semibold uppercase tracking-wider">
            Materiales & Planeación Cloud
          </p>
        </div>

        <form className="space-y-5" onSubmit={handleSubmit}>
          <div className="space-y-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
              Correo Electrónico
            </label>
            <input
              type="email"
              required
              className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-violet-500 focus:bg-white focus:ring-4 focus:ring-violet-500/10 transition-all"
              placeholder="nombre@berries.com"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
              Contraseña
            </label>
            <input
              type="password"
              required
              className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-violet-500 focus:bg-white focus:ring-4 focus:ring-violet-500/10 transition-all"
              placeholder="••••••••"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </div>

          {error && (
            <div className="text-red-600 font-semibold text-xs text-center bg-red-50 border border-red-100 p-3 rounded-lg">
              ⚠️ {error}
            </div>
          )}

          <button
            type="submit"
            disabled={submitting}
            className="w-full py-3.5 px-4 bg-gradient-to-r from-violet-600 to-rose-500 hover:from-violet-700 hover:to-rose-600 text-white font-bold rounded-xl shadow-lg shadow-violet-500/15 hover:shadow-violet-500/25 transition-all active:scale-[0.98] disabled:opacity-50 disabled:pointer-events-none cursor-pointer"
          >
            {submitting ? 'Iniciando sesión...' : 'Ingresar al Dashboard'}
          </button>
        </form>
      </div>

      <footer className="mt-8 text-[10px] text-slate-400 font-bold tracking-widest uppercase">
        Suite de Control de Materiales · Berries Paradise Cloud © 2026
      </footer>
    </main>
  );
}

export default function Page() {
  return (
    <AuthProvider>
      <LoginPage />
    </AuthProvider>
  );
}
