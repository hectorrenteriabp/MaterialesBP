"use client";

import { useState } from 'react';

export default function InventoryTab({ payload }) {
  const [search, setSearch] = useState('');
  const [filterAlm, setFilterAlm] = useState('TODOS');
  const [filterAbc, setFilterAbc] = useState('TODOS');

  const rawInv = payload.inv_master || payload.mrp_plan || []; // fallback to mrp_plan if inv_master isn't separate

  // Unique lists for dropdowns
  const uniqueAlmacenes = ['TODOS', ...new Set(rawInv.map(r => r.Almacen || r.alm).filter(Boolean))];
  const uniqueAbc = ['TODOS', 'A', 'B', 'C', 'X', 'S'];

  // Filtering
  const filteredInv = rawInv.filter(r => {
    const matchesSearch = 
      (r.Codigo || '').toLowerCase().includes(search.toLowerCase()) || 
      (r.Descripcion || '').toLowerCase().includes(search.toLowerCase());
    
    const matchesAlm = filterAlm === 'TODOS' || (r.Almacen || r.alm) === filterAlm;
    const matchesAbc = filterAbc === 'TODOS' || (r.ABC_Clase || r.ABC || r.Clase) === filterAbc;

    return matchesSearch && matchesAlm && matchesAbc;
  });

  const fnum = (v) => v !== undefined && v !== null ? parseInt(parseFloat(String(v)), 10).toLocaleString() : '—';
  const fval = (v) => v !== undefined && v !== null ? '$' + parseFloat(String(v)).toLocaleString('es-MX', { maximumFractionDigits: 0 }) : '—';

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-100 p-5 rounded-2xl shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-lg font-black text-slate-800 tracking-tight">📋 DETALLE MAESTRO DE INVENTARIOS</h2>
          <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-0.5">
            Consulta rápida de saldos valorizados, coberturas y clasificación ABC nacional.
          </p>
        </div>
      </div>

      {/* Filters */}
      <div className="bg-white border border-slate-100 p-4 rounded-2xl shadow-sm grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="space-y-1.5">
          <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Buscar SKU o Descripción</label>
          <input
            type="text"
            className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:outline-none focus:border-violet-500"
            placeholder="Ej: CL0349..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>

        <div className="space-y-1.5">
          <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Almacén</label>
          <select
            className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:outline-none focus:border-violet-500 font-bold text-slate-600"
            value={filterAlm}
            onChange={(e) => setFilterAlm(e.target.value)}
          >
            {uniqueAlmacenes.map(alm => <option key={alm} value={alm}>{alm}</option>)}
          </select>
        </div>

        <div className="space-y-1.5">
          <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Clase ABC</label>
          <select
            className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:outline-none focus:border-violet-500 font-bold text-slate-600"
            value={filterAbc}
            onChange={(e) => setFilterAbc(e.target.value)}
          >
            {uniqueAbc.map(abc => <option key={abc} value={abc}>{abc === 'TODOS' ? 'Todas las clases' : `Clase ${abc}`}</option>)}
          </select>
        </div>
      </div>

      {/* Inventory Table */}
      <div className="bg-white border border-slate-100 rounded-2xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100 text-slate-500 font-bold uppercase tracking-wider">
                <th className="p-3.5 pl-5">SKU / Material</th>
                <th className="p-3.5">Almacén</th>
                <th className="p-3.5 text-right">Stock</th>
                <th className="p-3.5">UM</th>
                <th className="p-3.5 text-right">Valorización (MXN)</th>
                <th className="p-3.5 text-center">Clase ABC</th>
                <th className="p-3.5 text-right pr-5">Días Cobertura</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredInv.length === 0 ? (
                <tr>
                  <td colSpan="7" className="text-center p-8 text-slate-400 font-semibold">
                    No se encontraron registros de inventario.
                  </td>
                </tr>
              ) : (
                filteredInv.map((r, i) => {
                  const abc = r.ABC_Clase || r.ABC || r.Clase || '—';
                  const isLowCob = parseFloat(r.Dias_cob) < 14;
                  
                  let abcBadgeClass = 'bg-slate-50 text-slate-500';
                  if (abc === 'A') abcBadgeClass = 'bg-emerald-50 text-emerald-700 border border-emerald-100';
                  else if (abc === 'B') abcBadgeClass = 'bg-amber-50 text-amber-700 border border-amber-100';
                  else if (abc === 'C') abcBadgeClass = 'bg-rose-50 text-rose-700 border border-rose-100';
                  
                  return (
                    <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                      <td className="p-3.5 pl-5 max-w-[250px]">
                        <div className="font-extrabold text-slate-800">{r.Codigo || r.sku}</div>
                        <div className="text-[10px] text-slate-400 truncate mt-0.5">{r.Descripcion || 'Sin descripción'}</div>
                      </td>
                      <td className="p-3.5 font-bold text-slate-600">{r.Almacen || r.alm}</td>
                      <td className="p-3.5 text-right font-bold text-slate-700">{fnum(r.Stock)}</td>
                      <td className="p-3.5 text-slate-400 font-bold uppercase tracking-wider">{r.UM || 'pzs'}</td>
                      <td className="p-3.5 text-right font-bold text-emerald-600">{fval(r.Costo_inv || r.Valor_deficit)}</td>
                      <td className="p-3.5 text-center">
                        <span className={`px-2 py-0.5 rounded text-[9px] font-black uppercase ${abcBadgeClass}`}>
                          Clase {abc}
                        </span>
                      </td>
                      <td className={`p-3.5 text-right font-black pr-5 ${isLowCob ? 'text-rose-600' : 'text-slate-600'}`}>
                        {r.Dias_cob ? `${Math.round(parseFloat(r.Dias_cob))} días` : '—'}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
