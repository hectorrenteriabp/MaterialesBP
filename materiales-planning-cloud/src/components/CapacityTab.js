"use client";

import { useState } from 'react';
import { supabase } from '@/lib/supabase';
import { useAuth } from '@/context/AuthContext';

export default function CapacityTab({ payload, dbCapacidades, onRefresh }) {
  const { user } = useAuth();
  const isEditable = user?.role === 'admin' || user?.role === 'almacen';
  
  const [editAlm, setEditAlm] = useState(null);
  const [editVal, setEditVal] = useState('');
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState({ text: '', type: '' });

  const rawCap = payload.cap_df || [];

  // Merge payload capacities with custom overrides from dbCapacidades
  const mergedCapacidades = rawCap.map(c => {
    const dbOver = dbCapacidades.find(x => x.alm === c.alm);
    const capacityActual = c.cap || 0;
    const capacityNew = (dbOver?.nueva_capacidad ?? dbOver?.capacidad_actual ?? c.cap) || null;
    const pctOcup = capacityNew > 0 ? (((c.pallets_used || 0) / capacityNew) * 100) : 0;
    
    return {
      ...c,
      cap_actual: capacityActual,
      cap_nueva: capacityNew,
      pct_ocup_new: pctOcup
    };
  });

  const handleEdit = (alm, currentVal) => {
    setEditAlm(alm);
    setEditVal(currentVal ? String(currentVal) : '');
  };

  const handleSave = async (alm) => {
    setSaving(true);
    setMsg({ text: '', type: '' });

    const numVal = parseInt(editVal.trim(), 10);
    if (isNaN(numVal) || numVal <= 0) {
      setMsg({ text: 'Por favor ingresa un número de capacidad válido (mayor a 0)', type: 'error' });
      setSaving(false);
      return;
    }

    try {
      const origCap = mergedCapacidades.find(x => x.alm === alm);
      
      // Update capabilities table in Supabase
      const { error } = await supabase
        .from('capacidades')
        .upsert({
          alm: alm,
          capacidad_actual: origCap?.cap_actual || numVal,
          nueva_capacidad: numVal,
          updated_at: new Date().toISOString()
        });

      if (error) throw error;

      // Log in audit log
      await supabase
        .from('audit_log')
        .insert({
          user_email: user.email,
          action: `Modificación de Capacidad: ${alm}`,
          details: `Cambio de capacidad de ${origCap?.cap_nueva || origCap?.cap_actual || 'N/A'} a ${numVal} pallets.`
        });

      setMsg({ text: `Capacidad de ${alm} actualizada exitosamente a ${numVal} pallets!`, type: 'success' });
      setEditAlm(null);
      if (onRefresh) onRefresh();
    } catch (e) {
      setMsg({ text: `Error al guardar cambios: ${e.message}`, type: 'error' });
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-100 p-5 rounded-2xl shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-lg font-black text-slate-800 tracking-tight">🏭 CONTROL DE CAPACIDAD DE ALMACENES</h2>
          <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-0.5">
            {isEditable 
              ? '🔓 Tienes permisos de edición para actualizar el límite físico de tarimas.' 
              : '🔒 Vista de solo lectura (Capacidades administrables por Almacén/Admin).'}
          </p>
        </div>
      </div>

      {msg.text && (
        <div className={`p-4 rounded-xl border font-bold text-xs ${msg.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-rose-50 border-rose-100 text-rose-800'}`}>
          {msg.type === 'success' ? '✅' : '⚠️'} {msg.text}
        </div>
      )}

      {/* Grid view of critical warehouse capacity */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {mergedCapacidades.map(c => {
          const isCrit = c.pct_ocup_new >= 90;
          const isWarn = c.pct_ocup_new >= 75 && c.pct_ocup_new < 90;
          const barColor = isCrit ? 'bg-rose-500' : isWarn ? 'bg-amber-500' : 'bg-emerald-500';
          const ringColor = isCrit ? 'border-rose-300' : isWarn ? 'border-amber-300' : 'border-emerald-300';
          
          return (
            <div 
              key={c.alm} 
              className={`bg-white border p-5 rounded-2xl shadow-sm hover:shadow-md transition-shadow relative overflow-hidden`}
              style={{ border: `1px solid ${isCrit ? '#fbd2d2' : isWarn ? '#f5e3b3' : '#e2e8f0'}` }}
            >
              {/* Subtle accent corner badge */}
              <div className={`absolute top-0 right-0 w-2 h-2 rounded-bl-xl ${isCrit ? 'bg-rose-500' : isWarn ? 'bg-amber-500' : 'bg-emerald-500'}`}></div>

              <div className="flex justify-between items-start mb-3">
                <div>
                  <h3 className="text-base font-black text-slate-800 tracking-tight">{c.alm}</h3>
                  <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">{c.nombre_sap || 'Almacén Red'}</p>
                </div>
                <div className={`px-2.5 py-1 rounded-full text-xs font-black ${isCrit ? 'bg-rose-50 text-rose-700' : isWarn ? 'bg-amber-50 text-amber-700' : 'bg-emerald-50 text-emerald-700'}`}>
                  {c.pct_ocup_new.toFixed(1)}% ocupado
                </div>
              </div>

              {/* Progress bar */}
              <div className="w-full bg-slate-100 rounded-full h-2.5 mb-4 overflow-hidden">
                <div className={`${barColor} h-2.5 rounded-full`} style={{ width: `${Math.min(c.pct_ocup_new, 100)}%` }}></div>
              </div>

              {/* Detailed metrics */}
              <div className="grid grid-cols-2 gap-4 text-xs font-semibold mb-4 bg-slate-50 p-3 rounded-xl">
                <div>
                  <div className="text-[9px] text-slate-400 font-bold uppercase tracking-wider">Tarimas Usadas</div>
                  <div className="text-slate-800 font-black mt-0.5">{(c.pallets_used || 0).toLocaleString()} <span className="text-[10px] font-normal text-slate-400">pos</span></div>
                </div>
                <div>
                  <div className="text-[9px] text-slate-400 font-bold uppercase tracking-wider">Capacidad Límite</div>
                  <div className="text-slate-800 font-black mt-0.5">{(c.cap_nueva || 0).toLocaleString()} <span className="text-[10px] font-normal text-slate-400">pos</span></div>
                </div>
              </div>

              {/* Action Button */}
              {isEditable && (
                <div className="pt-2 border-t border-slate-100 flex justify-end">
                  {editAlm === c.alm ? (
                    <div className="flex items-center gap-2 w-full">
                      <input
                        type="number"
                        className="bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs text-slate-900 focus:outline-none focus:border-violet-500 w-full"
                        placeholder="Nvo Límite"
                        value={editVal}
                        onChange={(e) => setEditVal(e.target.value)}
                        disabled={saving}
                      />
                      <button
                        onClick={() => handleSave(c.alm)}
                        disabled={saving}
                        className="bg-violet-600 hover:bg-violet-700 text-white font-bold text-xs px-3 py-1.5 rounded-lg cursor-pointer transition-colors whitespace-nowrap"
                      >
                        {saving ? '...' : 'Ok'}
                      </button>
                      <button
                        onClick={() => setEditAlm(null)}
                        disabled={saving}
                        className="bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold text-xs px-2.5 py-1.5 rounded-lg cursor-pointer transition-colors"
                      >
                        X
                      </button>
                    </div>
                  ) : (
                    <button
                      onClick={() => handleEdit(c.alm, c.cap_nueva)}
                      className="text-violet-600 hover:text-violet-800 font-bold text-xs cursor-pointer flex items-center gap-1.5 bg-violet-50 border border-violet-100 px-3 py-1.5 rounded-xl transition-all"
                    >
                      ✏️ Editar Límite
                    </button>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
