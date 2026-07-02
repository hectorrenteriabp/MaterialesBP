"use client";

export default function AuditLogTab({ logs }) {
  const fdate = (isoStr) => {
    try {
      const d = new Date(isoStr);
      return d.toLocaleDateString('es-MX', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch(e) {
      return isoStr;
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-100 p-5 rounded-2xl shadow-sm">
        <h2 className="text-lg font-black text-slate-800 tracking-tight">📜 BITÁCORA DE AUDITORÍA CLOUD</h2>
        <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-0.5">
          Historial de cambios realizados por compras, planeación, almacén y administradores.
        </p>
      </div>

      <div className="bg-white border border-slate-100 rounded-2xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100 text-slate-500 font-bold uppercase tracking-wider">
                <th className="p-3.5 pl-5">Fecha / Hora</th>
                <th className="p-3.5">Usuario</th>
                <th className="p-3.5">Acción</th>
                <th className="p-3.5 pr-5">Detalle del Cambio</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {logs.length === 0 ? (
                <tr>
                  <td colSpan="4" className="text-center p-8 text-slate-400 font-semibold">
                    No se han registrado movimientos de auditoría aún.
                  </td>
                </tr>
              ) : (
                logs.map((log) => (
                  <tr key={log.id} className="hover:bg-slate-50/50 transition-colors">
                    <td className="p-3.5 pl-5 font-bold text-slate-500 whitespace-nowrap">
                      {fdate(log.created_at)}
                    </td>
                    <td className="p-3.5">
                      <div className="font-extrabold text-slate-800">{log.user_email}</div>
                    </td>
                    <td className="p-3.5">
                      <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase bg-violet-50 text-violet-700 border border-violet-100">
                        {log.action}
                      </span>
                    </td>
                    <td className="p-3.5 text-slate-600 font-medium pr-5">
                      {log.details || '—'}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
