'use client';

export default function GlobalError({ error, reset }) {
  return (
    <html>
      <body className="flex flex-col items-center justify-center min-h-screen bg-slate-50 font-sans p-6 text-center">
        <h2 className="text-xl font-black text-slate-800 mb-2">Algo salió mal en el dashboard</h2>
        <p className="text-xs text-slate-400 mb-4">{error?.message || 'Error inesperado'}</p>
        <button
          onClick={() => reset()}
          className="bg-violet-600 hover:bg-violet-700 text-white font-bold text-xs px-4 py-2 rounded-xl cursor-pointer"
        >
          Reintentar
        </button>
      </body>
    </html>
  );
}
