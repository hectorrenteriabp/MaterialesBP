"use client";

import { useState } from 'react';
import { supabase } from '@/lib/supabase';
import { useAuth } from '@/context/AuthContext';

export default function MrpTab({ payload, dbTracking, onRefresh }) {
  const { user } = useAuth();
  const isEditable = user?.role === 'admin' || user?.role === 'compras' || user?.role === 'planeacion';

  const [search, setSearch] = useState('');
  const [filterAlm, setFilterAlm] = useState('TODOS');
  const [filterDecision, setFilterDecision] = useState('TODOS');
  const [onlyShorts, setOnlyShorts] = useState(false);

  const [editKey, setEditKey] = useState(null); // 'SKU|ALM'
  const [editStatus, setEditStatus] = useState('');
  const [editComments, setEditComments] = useState('');
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState({ text: '', type: '' });

  const rawMrp = payload.mrp_plan || [];

  // Merge payload MRP rows with database overrides
  const mergedMrp = rawMrp.map(r => {
    const key = `${r.Codigo}|${r.Almacen}`;
    const dbOver = dbTracking.find(x => x.sku_alm === key);
    
    return {
      ...r,
      sku_alm: key,
      Estatus_trk: dbOver?.status ?? r.Estatus_trk ?? '',
      Comentarios: dbOver?.comentarios ?? r.Comentarios ?? ''
    };
  });

  // Extract unique warehouses for filter dropdown
  const uniqueAlmacenes = ['TODOS', ...new Set(rawMrp.map(r => r.Almacen))];

  // Decision filters mapping
  const decisionFilterOptions = [
    { value: 'TODOS', label: 'Todas las decisiones' },
    { value: 'CORTOS', label: 'Solo Cortos / Déficit' },
    { value: 'REFILL', label: 'Solo Reabastos (Is_refill)' },
    { value: 'GLM', label: 'Decisión: GLM' },
    { value: 'TRASLADO', label: 'Decisión: Traslados' },
    { value: 'EXCESO', label: 'Sobre-stock / Exceso' }
  ];

  // Filtering logic
  const filteredMrp = mergedMrp.filter(r => {
    const matchesSearch = 
      r.Codigo.toLowerCase().includes(search.toLowerCase()) || 
      (r.Descripcion || '').toLowerCase().includes(search.toLowerCase());
    
    const matchesAlm = filterAlm === 'TODOS' || r.Almacen === filterAlm;

    let matchesDecision = true;
    if (filterDecision === 'CORTOS') {
      matchesDecision = parseFloat(r.Deficit) > 0 || r.Is_corto;
    } else if (filterDecision === 'REFILL') {
      matchesDecision = r.Is_refill;
    } else if (filterDecision === 'GLM') {
      matchesDecision = r.Dec_externa === 'GLM';
    } else if (filterDecision === 'TRASLADO') {
      matchesDecision = r.Dec_externa === 'TRASLADO';
    } else if (filterDecision === 'EXCESO') {
      matchesDecision = r.Is_overstock || r.Decision_mrp === 'Sobre-stock';
    }

    const matchesOnlyShorts = !onlyShorts || (parseFloat(r.Deficit) > 0 || r.Is_corto);

    return matchesSearch && matchesAlm && matchesDecision && matchesOnlyShorts;
  });

  const handleEditClick = (row) => {
    setEditKey(row.sku_alm);
    setEditStatus(row.Estatus_trk || '');
    setEditComments(row.Comentarios || '');
  };

  const handleSave = async (sku, alm) => {
    setSaving(true);
    setMsg({ text: '', type: '' });
    const key = `${sku}|${alm}`;

    try {
      // Upsert tracking override to Supabase mrp_tracking table
      const { error } = await supabase
        .from('mrp_tracking')
        .upsert({
          sku_alm: key,
          sku: sku,
          alm: alm,
          status: editStatus,
          comentarios: editComments,
          updated_at: new Date().toISOString()
        });

      if (error) throw error;

      // Log in audit log
      await supabase
        .from('audit_log')
        .insert({
          user_email: user.email,
          action: `Modificación de Tracking MRP: ${key}`,
          details: `Estatus: "${editStatus}", Comentarios: "${editComments}"`
        });

      setMsg({ text: `Tracking para ${key} guardado exitosamente en la nube!`, type: 'success' });
      setEditKey(null);
      if (onRefresh) onRefresh();
    } catch (e) {
      setMsg({ text: `Error al guardar tracking: ${e.message}`, type: 'error' });
    } finally {
      setSaving(false);
    }
  };

  const fnum = (v) => v !== undefined && v !== null ? parseInt(parseFloat(String(v)), 10).toLocaleString() : '—';
  const fval = (v) => v !== undefined && v !== null ? '$' + parseFloat(String(v)).toLocaleString('es-MX', { maximumFractionDigits: 0 }) : '—';

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-100 p-5 rounded-2xl shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-lg font-black text-slate-800 tracking-tight">📦 MÓDULO MRP & TRACKING DE MATERIALES</h2>
          <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-0.5">
            {isEditable 
              ? '🔓 Tienes permisos para actualizar comentarios y estatus del seguimiento.' 
              : '🔒 Vista de solo lectura (Estatus y comentarios editables por Compras/Planeación/Admin).'}
          </p>
        </div>
      </div>

      {msg.text && (
        <div className={`p-4 rounded-xl border font-bold text-xs ${msg.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-rose-50 border-rose-100 text-rose-800'}`}>
          {msg.type === 'success' ? '✅' : '⚠️'} {msg.text}
        </div>
      )}

      {/* Filters Toolbar */}
      <div className="bg-white border border-slate-100 p-4 rounded-2xl shadow-sm grid grid-cols-1 md:grid-cols-4 gap-4">
        {/* Search */}
        <div className="space-y-1.5">
          <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Buscar SKU o Descripción</label>
          <input
            type="text"
            className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:outline-none focus:border-violet-500"
            placeholder="Ej: CL0269..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>

        {/* Warehouse Filter */}
        <div className="space-y-1.5">
          <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Almacén destino</label>
          <select
            className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:outline-none focus:border-violet-500 font-bold text-slate-600"
            value={filterAlm}
            onChange={(e) => setFilterAlm(e.target.value)}
          >
            {uniqueAlmacenes.map(alm => <option key={alm} value={alm}>{alm}</option>)}
          </select>
        </div>

        {/* Decision Filter */}
        <div className="space-y-1.5">
          <label className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Decisión / Regla MRP</label>
          <select
            className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2 text-xs focus:outline-none focus:border-violet-500 font-bold text-slate-600"
            value={filterDecision}
            onChange={(e) => setFilterDecision(e.target.value)}
          >
            {decisionFilterOptions.map(opt => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
          </select>
        </div>

        {/* Checkbox only shorts */}
        <div className="flex items-center md:pt-5 pl-1.5">
          <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-600">
            <input
              type="checkbox"
              className="w-4 h-4 rounded text-violet-600 border-slate-300 focus:ring-violet-500"
              checked={onlyShorts}
              onChange={(e) => setOnlyShorts(e.target.checked)}
            />
            ⚠️ Mostrar solo Cortos/Déficit
          </label>
        </div>
      </div>

      {/* MRP Interactive Table */}
      <div className="bg-white border border-slate-100 rounded-2xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100 text-slate-500 font-bold uppercase tracking-wider">
                <th className="p-3.5 pl-5">SKU / Descripción</th>
                <th className="p-3.5">Almacén</th>
                <th className="p-3.5 text-right">Stock</th>
                <th className="p-3.5 text-right">Mín</th>
                <th className="p-3.5 text-right">Déficit</th>
                <th className="p-3.5 text-center">Decisión MRP</th>
                <th className="p-3.5">Estatus Tracking</th>
                <th className="p-3.5">Comentarios</th>
                <th className="p-3.5 text-center pr-5">Acción</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredMrp.length === 0 ? (
                <tr>
                  <td colSpan="9" className="text-center p-8 text-slate-400 font-semibold">
                    No se encontraron registros MRP con los filtros seleccionados.
                  </td>
                </tr>
              ) : (
                filteredMrp.map(r => {
                  const isCorto = parseFloat(r.Deficit) > 0 || r.Is_corto;
                  const isExceso = r.Is_overstock || r.Decision_mrp === 'Sobre-stock';
                  
                  let badgeClass = 'bg-slate-50 text-slate-500 border border-slate-100';
                  if (isCorto) badgeClass = 'bg-red-50 text-red-700 border border-red-100';
                  else if (isExceso) badgeClass = 'bg-amber-50 text-amber-700 border border-amber-100';
                  else if (r.Dec_externa === 'TRASLADO') badgeClass = 'bg-cyan-50 text-cyan-700 border border-cyan-100';
                  else if (r.Dec_externa === 'GLM') badgeClass = 'bg-purple-50 text-purple-700 border border-purple-100';
                  else badgeClass = 'bg-emerald-50 text-emerald-700 border border-emerald-100';

                  const decisionLabel = isCorto 
                    ? `⚠️ CORTO (${fnum(r.Deficit)} pzs)` 
                    : isExceso 
                      ? `SOBRE-STOCK` 
                      : r.Dec_externa === 'TRASLADO' 
                        ? `🚚 TRASLADO` 
                        : r.Dec_externa === 'GLM' 
                          ? `📦 GLM` 
                          : `🟢 OK`;

                  return (
                    <tr key={r.sku_alm} className="hover:bg-slate-50/50 transition-colors">
                      <td className="p-3.5 pl-5 max-w-[200px]">
                        <div className="font-extrabold text-slate-800">{r.Codigo}</div>
                        <div className="text-[10px] text-slate-400 truncate mt-0.5">{r.Descripcion || 'Sin descripción'}</div>
                      </td>
                      <td className="p-3.5 font-bold text-slate-600">{r.Almacen}</td>
                      <td className="p-3.5 text-right font-bold text-slate-700">{fnum(r.Stock)}</td>
                      <td className="p-3.5 text-right font-medium text-slate-400">{fnum(r.Min)}</td>
                      <td className={`p-3.5 text-right font-black ${isCorto ? 'text-red-600' : 'text-slate-400'}`}>
                        {isCorto ? fnum(r.Deficit) : '—'}
                      </td>
                      <td className="p-3.5 text-center">
                        <span className={`px-2 py-1 rounded-full text-[9px] font-black ${badgeClass}`}>
                          {decisionLabel}
                        </span>
                      </td>
                      <td className="p-3.5">
                        {editKey === r.sku_alm ? (
                          <select
                            className="bg-slate-50 border border-slate-200 rounded px-1.5 py-0.5 text-xs focus:outline-none focus:border-violet-500 font-bold"
                            value={editStatus}
                            onChange={(e) => setEditStatus(e.target.value)}
                          >
                            <option value="">(Selecciona)</option>
                            <option value="Pendiente">Pendiente</option>
                            <option value="En proceso">En proceso</option>
                            <option value="Ordenado">Ordenado</option>
                            <option value="En tránsito">En tránsito</option>
                            <option value="Entregado">Entregado</option>
                            <option value="Cancelado">Cancelado</option>
                          </select>
                        ) : (
                          <span className={`font-bold text-[10px] ${r.Estatus_trk ? 'text-violet-600' : 'text-slate-400'}`}>
                            {r.Estatus_trk || 'Sin estatus'}
                          </span>
                        )}
                      </td>
                      <td className="p-3.5 max-w-[200px] truncate">
                        {editKey === r.sku_alm ? (
                          <input
                            type="text"
                            className="w-full bg-slate-50 border border-slate-200 rounded px-2 py-0.5 text-xs focus:outline-none focus:border-violet-500"
                            value={editComments}
                            onChange={(e) => setEditComments(e.target.value)}
                            placeholder="Notas de compra..."
                          />
                        ) : (
                          <span className="text-slate-500 text-[11px]">{r.Comentarios || '—'}</span>
                        )}
                      </td>
                      <td className="p-3.5 text-center pr-5">
                        {isEditable && (
                          editKey === r.sku_alm ? (
                            <div className="flex justify-center gap-1">
                              <button
                                onClick={() => handleSave(r.Codigo, r.Almacen)}
                                disabled={saving}
                                className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-2 py-1 rounded cursor-pointer"
                              >
                                {saving ? '...' : '✓'}
                              </button>
                              <button
                                onClick={() => setEditKey(null)}
                                disabled={saving}
                                className="bg-slate-100 hover:bg-slate-200 text-slate-500 font-bold px-2 py-1 rounded cursor-pointer"
                              >
                                ✕
                              </button>
                            </div>
                          ) : (
                            <button
                              onClick={() => handleEditClick(r)}
                              className="text-violet-600 hover:text-violet-800 font-bold cursor-pointer"
                            >
                              ✏️
                            </button>
                          )
                        )}
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
