"use client";

import { useState } from 'react';
import { supabase } from '@/lib/supabase';
import { useAuth } from '@/context/AuthContext';

export default function AdminConfigTab({ permissions, onRefresh }) {
  const { user } = useAuth();
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState({ text: '', type: '' });

  const roles = [
    { key: 'admin', name: 'Administrador' },
    { key: 'compras', name: 'Compras' },
    { key: 'planeacion', name: 'Planeación' },
    { key: 'almacen', name: 'Almacén' }
  ];

  const tabs = [
    { id: 'eje', name: '📊 Ejecutivo (Dashboards/Gauges)' },
    { id: 'ver', name: '📋 Inventarios (Detalle Maestro)' },
    { id: 'cap', name: '🏭 Capacidad (Warehouse Capacity)' },
    { id: 'mrp', name: '📦 MRP & Planeación (Interactive Table)' },
    { id: 'audit', name: '📜 Bitácora (Audit Log)' }
  ];

  const handleToggle = async (roleKey, tabId) => {
    if (saving) return;
    setSaving(true);
    setMsg({ text: '', type: '' });

    const rolePerm = permissions.find(p => p.role === roleKey);
    let allowedTabs = rolePerm?.allowed_tabs || [];

    if (allowedTabs.includes(tabId)) {
      // Prevent admin from locking themselves out of the config tab
      if (roleKey === 'admin' && tabId === 'config') {
        setMsg({ text: 'No puedes quitarle el permiso de Configuración al Administrador.', type: 'error' });
        setSaving(false);
        return;
      }
      allowedTabs = allowedTabs.filter(id => id !== tabId);
    } else {
      allowedTabs = [...allowedTabs, tabId];
    }

    try {
      const { error } = await supabase
        .from('role_permissions')
        .upsert({
          role: roleKey,
          allowed_tabs: allowedTabs,
          updated_at: new Date().toISOString()
        });

      if (error) throw error;

      // Log in audit log
      await supabase
        .from('audit_log')
        .insert({
          user_email: user.email,
          action: `Permisos Actualizados: ${roleKey}`,
          details: `Pestañas permitidas actualizadas a: [${allowedTabs.join(', ')}]`
        });

      setMsg({ text: `Permisos del rol ${roleKey} actualizados correctamente!`, type: 'success' });
      if (onRefresh) await onRefresh();
    } catch (e) {
      setMsg({ text: `Error al guardar permisos: ${e.message}`, type: 'error' });
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-100 p-5 rounded-2xl shadow-sm">
        <h2 className="text-lg font-black text-slate-800 tracking-tight">⚙️ CONFIGURACIÓN DE ACCESOS Y ROLES</h2>
        <p className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-0.5">
          Matriz de Visibilidad: Configura qué pestañas puede visualizar cada perfil de usuario en caliente.
        </p>
      </div>

      {msg.text && (
        <div className={`p-4 rounded-xl border font-bold text-xs ${msg.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-rose-50 border-rose-100 text-rose-800'}`}>
          {msg.type === 'success' ? '✅' : '⚠️'} {msg.text}
        </div>
      )}

      <div className="bg-white border border-slate-100 rounded-2xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100 text-slate-500 font-bold uppercase tracking-wider">
                <th className="p-4 pl-6">Pestaña / Módulo</th>
                {roles.map(r => (
                  <th key={r.key} className="p-4 text-center">{r.name}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {tabs.map(tab => (
                <tr key={tab.id} className="hover:bg-slate-50/50 transition-colors">
                  <td className="p-4 pl-6 font-extrabold text-slate-800">{tab.name}</td>
                  {roles.map(role => {
                    const rolePerm = permissions.find(p => p.role === role.key);
                    const isChecked = rolePerm?.allowed_tabs?.includes(tab.id) ?? true;
                    const isDisabled = role.key === 'admin' && tab.id === 'config';
                    
                    return (
                      <td key={role.key} className="p-4 text-center">
                        <input
                          type="checkbox"
                          className="w-4 h-4 rounded text-violet-600 border-slate-300 focus:ring-violet-500 cursor-pointer disabled:opacity-30 disabled:pointer-events-none mx-auto"
                          checked={isChecked}
                          onChange={() => handleToggle(role.key, tab.id)}
                          disabled={saving || isDisabled}
                        />
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
