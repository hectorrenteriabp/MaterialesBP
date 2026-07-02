"use client";

import { useEffect, useState } from 'react';
import { useAuth, AuthProvider } from '@/context/AuthContext';
import { supabase } from '@/lib/supabase';
import ExecutiveTab from '@/components/ExecutiveTab';
import InventoryTab from '@/components/InventoryTab';
import CapacityTab from '@/components/CapacityTab';
import MrpTab from '@/components/MrpTab';
import AuditLogTab from '@/components/AuditLogTab';
import AdminConfigTab from '@/components/AdminConfigTab';

function DashboardHome() {
  const { user, logout, loading: authLoading } = useAuth();
  
  const [activeTab, setActiveTab] = useState('eje');
  const [loadingData, setLoadingData] = useState(true);
  const [payload, setPayload] = useState(null);
  
  // Database live tables
  const [dbCapacidades, setDbCapacidades] = useState([]);
  const [dbTracking, setDbTracking] = useState([]);
  const [dbPermissions, setDbPermissions] = useState([]);
  const [dbLogs, setDbLogs] = useState([]);

  // Fetch all cloud data
  const fetchData = async () => {
    try {
      // 1. Fetch dashboard consolidated payload (id = 1)
      const { data: payData } = await supabase
        .from('dashboard_payload')
        .select('payload')
        .eq('id', 1)
        .single();
      
      if (payData && payData.payload) {
        setPayload(payData.payload);
      }

      // 2. Fetch capacities
      const { data: capData } = await supabase
        .from('capacidades')
        .select('*');
      setDbCapacidades(capData || []);

      // 3. Fetch tracking overrides
      const { data: trData } = await supabase
        .from('mrp_tracking')
        .select('*');
      setDbTracking(trData || []);

      // 4. Fetch permissions matrix
      const { data: permData } = await supabase
        .from('role_permissions')
        .select('*');
      setDbPermissions(permData || []);

      // 5. Fetch audit logs (last 100 entries)
      const { data: logData } = await supabase
        .from('audit_log')
        .select('*')
        .order('created_at', { ascending: false })
        .limit(100);
      setDbLogs(logData || []);

    } catch (e) {
      console.error("Error loading data from Supabase:", e);
    } finally {
      setLoadingData(false);
    }
  };

  useEffect(() => {
    if (user) {
      fetchData();
    }
  }, [user]);

  if (authLoading || (loadingData && !payload)) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-slate-50">
        <div className="relative flex items-center justify-center w-12 h-12 mb-4">
          <div className="absolute w-12 h-12 rounded-full border-4 border-violet-100"></div>
          <div className="absolute w-12 h-12 rounded-full border-4 border-t-violet-600 animate-spin"></div>
        </div>
        <p className="text-xs font-black tracking-widest text-slate-400 uppercase">
          Estableciendo conexión segura con Supabase...
        </p>
      </div>
    );
  }

  // Filter allowed tabs based on user role permissions
  const rolePerm = dbPermissions.find(p => p.role === user?.role);
  // Default tabs fallback if DB is not ready or configured yet
  const allowedTabs = rolePerm?.allowed_tabs || ['eje', 'ver', 'cap', 'mrp', 'audit'];

  const allTabs = [
    { id: 'eje', name: '📊 Ejecutivo', component: <ExecutiveTab payload={payload} /> },
    { id: 'ver', name: '📋 Inventarios', component: <InventoryTab payload={payload} /> },
    { id: 'cap', name: '🏭 Capacidad', component: <CapacityTab payload={payload} dbCapacidades={dbCapacidades} onRefresh={fetchData} /> },
    { id: 'mrp', name: '📦 MRP & Tracking', component: <MrpTab payload={payload} dbTracking={dbTracking} onRefresh={fetchData} /> },
    { id: 'audit', name: '📜 Bitácora', component: <AuditLogTab logs={dbLogs} /> },
    { id: 'config', name: '⚙️ Configuración', component: <AdminConfigTab permissions={dbPermissions} onRefresh={fetchData} /> }
  ];

  // Only render tabs that are allowed for the user's role
  // Always allow admin to access config, and verify others
  const visibleTabs = allTabs.filter(tab => {
    if (tab.id === 'config') return user?.role === 'admin';
    return allowedTabs.includes(tab.id);
  });

  // Ensure current active tab is visible, otherwise fallback to first visible tab
  const activeTabDetails = visibleTabs.find(t => t.id === activeTab) || visibleTabs[0] || allTabs[0];

  return (
    <div className="min-h-screen flex flex-col bg-slate-50">
      {/* Top Navbar */}
      <header className="bg-white border-b border-slate-100 shadow-sm sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <span className="text-xl">🍇</span>
            <div>
              <h1 className="text-sm font-black text-slate-800 tracking-tight leading-none">
                Berries <span className="text-violet-600">Paradise</span>
              </h1>
              <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mt-1 block">
                Materiales & Planeación Cloud
              </span>
            </div>
          </div>

          <div className="flex items-center gap-4">
            {/* User Profile Badge */}
            <div className="flex items-center gap-2.5 bg-slate-50 border border-slate-100 px-3 py-1.5 rounded-xl">
              <div className="w-6.5 h-6.5 bg-gradient-to-br from-violet-600 to-rose-500 rounded-lg text-white font-extrabold text-[10px] flex items-center justify-center uppercase">
                {user?.name?.substring(0, 2)}
              </div>
              <div className="hidden sm:block text-left">
                <div className="text-[10px] font-black text-slate-800 leading-none">{user?.name}</div>
                <div className="text-[8px] font-bold text-violet-600 uppercase tracking-wider mt-0.5">{user?.role}</div>
              </div>
            </div>

            {/* Logout button */}
            <button
              onClick={logout}
              className="text-[10px] font-bold text-rose-600 bg-rose-50 hover:bg-rose-100 border border-rose-100 px-3 py-1.5 rounded-xl cursor-pointer transition-colors"
            >
              Cerrar Sesión
            </button>
          </div>
        </div>
      </header>

      {/* Main Dashboard Layout */}
      <div className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 flex flex-col gap-6">
        
        {/* Navigation Tabs Bar */}
        <nav className="flex gap-2 p-1.5 bg-slate-200/50 border border-slate-200/60 rounded-2xl self-start overflow-x-auto max-w-full">
          {visibleTabs.map(tab => {
            const isSelected = activeTabDetails.id === tab.id;
            return (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`px-4 py-2 rounded-xl text-xs font-black transition-all cursor-pointer whitespace-nowrap ${
                  isSelected 
                    ? 'bg-white text-slate-900 shadow-sm' 
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-100/50'
                }`}
              >
                {tab.name}
              </button>
            );
          })}
        </nav>

        {/* Tab View Container */}
        <main className="flex-1">
          {activeTabDetails.component}
        </main>
      </div>

      <footer className="bg-white border-t border-slate-100 py-6 text-center text-[9px] text-slate-400 font-bold uppercase tracking-widest">
        Suite de Control de Materiales · Berries Paradise Cloud © 2026
      </footer>
    </div>
  );
}

export default function Page() {
  return (
    <AuthProvider>
      <DashboardHome />
    </AuthProvider>
  );
}
