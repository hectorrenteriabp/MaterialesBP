"use client";

import { useEffect, useRef, useState } from 'react';
import { Chart, registerables } from 'chart.js';

Chart.register(...registerables);

export default function ExecutiveTab({ payload }) {
  const donutRef = useRef(null);
  const barRef = useRef(null);
  const trendRef = useRef(null);
  const abcRef = useRef(null);
  
  const [trendDesde, setTrendDesde] = useState('2025-08');
  const [trendHasta, setTrendHasta] = useState('');
  const trendChartInst = useRef(null);

  const R = payload.mrp_resumen || {};
  const cd = payload.cap_df || [];
  const capUsed = cd.reduce((a, r) => a + (r.pallets_used || 0), 0);
  const capTotal = cd.reduce((a, r) => a + (r.cap || 0), 0);
  const capPct = capTotal > 0 ? ((capUsed / capTotal) * 100).toFixed(1) : (payload.total_cap_pct || 0);
  const vPct = payload.ver_piezas || 0;
  const cobPct = R.pct_cob_1sem || 0;
  const cob2Pct = R.pct_cob_2sem || 0;
  const nCortos = R.n_cortos_plan || 0;
  const nCortosCriticos = R.n_cortos_criticos || 0;
  const nConStock = R.n_cortos_con_stock || 0;
  const nTotal = R.n_total || 0;
  const semsP = R.sems_cob_general || 0;
  const costInv = payload.total_costo_inv || 0;
  const nSkus = payload.total_inv_codes || 0;
  const genDate = payload.gen_date || '—';

  const ML = {
    '2025-08': 'Ago 25', '2025-09': 'Sep 25', '2025-10': 'Oct 25', '2025-11': 'Nov 25', '2025-12': 'Dic 25',
    '2026-01': 'Ene 26', '2026-02': 'Feb 26', '2026-03': 'Mar 26', '2026-04': 'Abr 26', '2026-05': 'May 26', '2026-06': 'Jun 26'
  };

  const kpiColor = (v, hi, mid) => v >= hi ? '#059669' : v >= mid ? '#d97706' : '#dc2626';
  const fmxn = (n) => '$' + n.toLocaleString('es-MX', { maximumFractionDigits: 0 });
  const fmxK = (v) => v >= 1e6 ? (v / 1e6).toFixed(1) + 'M' : v >= 1e3 ? (v / 1e3).toFixed(0) + 'K' : v;

  // ── Card styling helper (matches premium outlines requested by user) ──
  const cardStyle = (color) => {
    let bd = 'var(--border)';
    let sh = 'rgba(109, 40, 217, 0.02)';
    if (color === '#059669') { bd = '#b8ecd5'; sh = 'rgba(5, 150, 105, 0.04)'; }
    else if (color === '#d97706' || color === '#b45309' || color === '#ea580c') { bd = '#f5e3b3'; sh = 'rgba(217, 119, 6, 0.04)'; }
    else if (color === '#dc2626') { bd = '#fbd2d2'; sh = 'rgba(220, 38, 38, 0.04)'; }
    else if (color === '#2563eb' || color === '#0891b2') { bd = '#cdddfb'; sh = 'rgba(37, 99, 235, 0.04)'; }
    else if (color === '#0e7490') { bd = '#b6ebf2'; sh = 'rgba(14, 116, 144, 0.04)'; }
    return {
      background: 'white',
      border: `1px solid ${bd}`,
      borderRadius: '12px',
      boxShadow: `0 4px 14px ${sh}, 0 1px 2px rgba(0,0,0,0.05)`,
      padding: '14px',
      transition: 'all 0.2s ease'
    };
  };

  // SVG Gauge Component
  function SvgGauge({ pct, color, icon, label, sub }) {
    const r = 42, cx = 54, cy = 54, circ = 2 * Math.PI * r;
    const dash = Math.min(1, pct / 100) * circ;
    const gap = circ - dash;
    return (
      <div className="text-center p-2">
        <svg width="108" height="108" viewBox="0 0 108 108" className="mx-auto">
          <circle cx={cx} cy={cy} r={r} fill="none" stroke="#f1f5f9" strokeWidth="9" />
          <circle
            cx={cx} cy={cy} r={r} fill="none" stroke={color} strokeWidth="9"
            strokeDasharray={`${dash.toFixed(1)} ${gap.toFixed(1)}`}
            strokeLinecap="round" transform={`rotate(-90 ${cx} ${cy})`}
            style={{ transition: 'stroke-dasharray .6s ease' }}
          />
          <text x={cx} y={cy - 8} textAnchor="middle" fontSize="12" fill="#64748b">{icon}</text>
          <text x={cx} y={cy + 8} textAnchor="middle" fontSize="14" fontWeight="900" fill={color}>{pct}%</text>
        </svg>
        <div className="font-bold text-xs text-slate-700 mt-1">{label}</div>
        <div className="text-[10px] text-slate-400 mt-0.5">{sub}</div>
      </div>
    );
  }

  // ── ABC data ──
  const abcSum = payload.abc_summary || [];
  const _noMov = payload.no_mov || [];
  const _noMovMonto = _noMov.reduce((a, r) => a + (r.monto_mxn || 0), 0);
  const _abcBase = abcSum.map(r => ({
    cls: r.clase || r.Clase || '',
    skus: r.n_skus || 0,
    pctMovs: r.pct_movs || 0,
    monto: r.monto_mxn || 0
  }));
  const abcAll = _abcBase.concat(_noMov.length > 0 ? [{ cls: 'X', skus: _noMov.length, pctMovs: 0, monto: _noMovMonto }] : []);
  const abcLabels = abcAll.map(r => r.cls);
  const abcSkus = abcAll.map(r => r.skus);
  const abcPctMovs = abcAll.map(r => r.pctMovs);

  // ── MRP decisions ──
  const mrpLabels = ['Cubierto', 'Traslado', 'GLM', 'Exceso (>Max)', 'Corto'];
  const mrpVals = [R.n_cubierto || 0, R.n_traslado || 0, R.n_glm || 0, R.n_exceso || 0, R.n_corto || 0];
  const mrpDescs = [
    'Entre MIN y MAX — correcto', 'Falta stock, disponible en zona',
    'Falta stock, cubierto por GLM', 'Sobre-stock, frenar compras',
    'Sin stock ni solución de zona'
  ];
  const mrpColors = ['#059669', '#0891b2', '#d97706', '#ea580c', '#dc2626'];

  // ── Top almacenes capacity ──
  const capTop = [...cd].sort((a, b) => (b.pct_ocup || 0) - (a.pct_ocup || 0)).slice(0, 8);
  const capAlmLbls = capTop.map(r => r.alm);
  const capAlmVals = capTop.map(r => Math.min(r.pct_ocup || 0, 150));
  const capAlmColors = capAlmVals.map(v => v >= 100 ? '#dc2626' : v >= 85 ? '#d97706' : '#059669');

  const meses = payload.meses || [];

  // Update Line Chart dynamically
  useEffect(() => {
    if (!trendRef.current || !meses.length) return;
    const desde = trendDesde;
    const hasta = trendHasta || meses[meses.length - 1];
    
    const i0 = Math.max(0, meses.indexOf(desde));
    let i1 = meses.indexOf(hasta);
    if (i1 < i0) i1 = meses.length - 1;
    const mF = meses.slice(i0, i1 + 1);
    const labF = mF.map(m => ML[m] || m);
    
    const td = payload.trend?.['TODOS'] || { E: [], S: [] };
    const idxMap = meses.reduce((o, m, i) => { o[m] = i; return o; }, {});
    const eF = mF.map(m => td.E[idxMap[m]] || 0);
    const sF = mF.map(m => td.S[idxMap[m]] || 0);

    if (trendChartInst.current) {
      trendChartInst.current.destroy();
    }

    const ctx = trendRef.current.getContext('2d');
    trendChartInst.current = new Chart(ctx, {
      type: 'line',
      data: {
        labels: labF,
        datasets: [
          { label: 'Entradas', data: eF, borderColor: '#059669', backgroundColor: '#ecfdf5', borderWidth: 2, fill: true, tension: 0.35, pointRadius: 3 },
          { label: 'Salidas', data: sF, borderColor: '#dc2626', backgroundColor: '#fef2f2', borderWidth: 2, fill: true, tension: 0.35, pointRadius: 3 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 10 } } },
          y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', font: { size: 10 }, callback: (v) => fmxK(v) } }
        },
        plugins: {
          legend: { labels: { color: '#64748b', font: { size: 10 }, boxWidth: 12 } }
        }
      }
    });

    return () => {
      if (trendChartInst.current) trendChartInst.current.destroy();
    };
  }, [trendDesde, trendHasta, payload]);

  // Static Charts initialization
  useEffect(() => {
    // 1. Donut
    let donutInst = null;
    if (donutRef.current) {
      const ctx = donutRef.current.getContext('2d');
      donutInst = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: mrpLabels,
          datasets: [{ data: mrpVals, backgroundColor: mrpColors, borderColor: 'white', borderWidth: 2 }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: '70%',
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (c) => ` ${c.label}: ${c.raw} SKU×Alm`,
                afterLabel: (c) => `  ${mrpDescs[c.dataIndex]}`
              }
            }
          }
        }
      });
    }

    // 2. Bar Ocupacion
    let barInst = null;
    if (barRef.current) {
      const ctx = barRef.current.getContext('2d');
      barInst = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: capAlmLbls,
          datasets: [{ label: 'Ocupación %', data: capAlmVals, backgroundColor: capAlmColors, borderRadius: 4 }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b' }, max: 150, min: 0 },
            y: { grid: { display: false }, ticks: { color: '#64748b', font: { size: 11 } } }
          },
          plugins: { legend: { display: false } }
        }
      });
    }

    // 3. ABC Bar
    let abcInst = null;
    if (abcRef.current) {
      const ctx = abcRef.current.getContext('2d');
      const abcClrs = { 'A': '#059669', 'B': '#d97706', 'C': '#dc2626', 'X': '#6b7280', 'S': '#7c3aed' };
      const colors = abcLabels.map(l => abcClrs[l] || '#2563eb');
      abcInst = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: abcLabels,
          datasets: [{ label: '% Movimientos', data: abcPctMovs, backgroundColor: colors, borderRadius: 5 }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: { grid: { display: false }, ticks: { color: '#64748b', font: { size: 12, weight: 'bold' } } },
            y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b' }, max: 100 }
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (c) => ` ${c.raw}% movs (${abcSkus[c.dataIndex]} SKUs)`
              }
            }
          }
        }
      });
    }

    return () => {
      if (donutInst) donutInst.destroy();
      if (barInst) barInst.destroy();
      if (abcInst) abcInst.destroy();
    };
  }, [payload]);

  return (
    <div className="space-y-6">
      {/* KPI Header Card */}
      <div className="relative overflow-hidden bg-white border border-slate-100 p-6 rounded-2xl shadow-sm">
        <div className="absolute inset-0 bg-gradient-to-r from-violet-500/5 to-rose-500/5 opacity-40 pointer-events-none"></div>
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 relative z-10">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-1.5 h-6 bg-gradient-to-b from-violet-600 to-rose-500 rounded-full"></div>
              <h2 className="text-xl font-black text-slate-800 tracking-tight">RESUMEN EJECUTIVO DE INVENTARIOS</h2>
            </div>
            <p className="text-xs text-slate-400 font-bold tracking-wider mt-1 uppercase">
              Actualizado: {genDate} · {nSkus} SKUs con stock · {nTotal} SKU×Alm en plan MRP
            </p>
          </div>
        </div>
      </div>

      {/* SVG Gauges Grid */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div style={cardStyle(kpiColor(vPct, 99, 95))}>
          <SvgGauge pct={vPct} color={kpiColor(vPct, 99, 95)} icon="🎯" label="Veracidad" sub="por piezas" />
        </div>
        <div style={cardStyle(kpiColor(100 - parseFloat(capPct), 15, 0))}>
          <SvgGauge pct={parseFloat(capPct)} color={kpiColor(100 - parseFloat(capPct), 15, 0)} icon="🏭" label="Capacidad" sub={`${capUsed.toLocaleString()} / ${capTotal.toLocaleString()} pal`} />
        </div>
        <div style={cardStyle(kpiColor(cobPct, 95, 80))}>
          <SvgGauge pct={cobPct} color={kpiColor(cobPct, 95, 80)} icon="✅" label="Sin Cortos" sub={`${nCortosCriticos} crít · ${nConStock} buffer`} />
        </div>
        <div style={cardStyle(kpiColor(cob2Pct, 80, 60))}>
          <SvgGauge pct={cob2Pct} color={kpiColor(cob2Pct, 80, 60)} icon="📦" label="Cobertura ≥2sem" sub="% vs objetivo 2sem" />
        </div>
        <div style={cardStyle(kpiColor(semsP, 2, 1))} className="flex flex-col justify-center items-center text-center">
          <div className="text-lg">⏱️</div>
          <div className="text-3xl font-black mt-1" style={{ color: kpiColor(semsP, 2, 1) }}>{semsP}</div>
          <div className="text-xs font-bold text-slate-600 mt-1">Sems Cob. Real</div>
          <div className="text-[10px] text-slate-400 mt-0.5">Stock ÷ Demanda Nac.</div>
          <div className="mt-3 text-sm font-extrabold text-violet-600">{fmxn(costInv)}</div>
          <div className="text-[9px] text-slate-400">Valor Inventario</div>
        </div>
      </div>

      {/* Charts Row 1 */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {/* Donut Chart */}
        <div style={cardStyle('#6d28d9')} className="flex flex-col">
          <h3 className="text-xs font-extrabold text-violet-700 uppercase tracking-wider mb-4">🔧 Decisiones MRP</h3>
          <div className="relative h-44 flex-1">
            <canvas ref={donutRef}></canvas>
          </div>
          <div className="mt-4 space-y-2">
            {mrpLabels.map((lb, i) => mrpVals[i] > 0 && (
              <div key={lb} className="flex justify-between items-center text-[10px]">
                <span className="font-semibold" style={{ color: mrpColors[i] }}>● {lb}</span>
                <span className="text-slate-400 flex-1 text-right mr-3 truncate max-w-[140px]">{mrpDescs[i]}</span>
                <span className="font-bold text-slate-700">{mrpVals[i]}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Capacity Bar Chart */}
        <div style={cardStyle('#2563eb')} className="flex flex-col col-span-1 md:col-span-1">
          <h3 className="text-xs font-extrabold text-violet-700 uppercase tracking-wider mb-4">🏭 Ocupación por Almacén (Top 8)</h3>
          <div className="relative h-64 flex-1">
            <canvas ref={barRef}></canvas>
          </div>
        </div>

        {/* Trend Line Chart */}
        <div style={cardStyle('#0e7490')} className="flex flex-col col-span-1">
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-xs font-extrabold text-violet-700 uppercase tracking-wider">📈 Tendencia de Movimientos</h3>
            <div className="flex items-center gap-1.5 text-[10px] text-slate-400">
              <span>De</span>
              <select
                value={trendDesde}
                onChange={(e) => setTrendDesde(e.target.value)}
                className="bg-slate-50 border border-slate-200 rounded px-1.5 py-0.5 text-violet-600 font-bold focus:outline-none"
              >
                {meses.map(m => <option key={m} value={m}>{ML[m] || m}</option>)}
              </select>
            </div>
          </div>
          <div className="relative h-64 flex-1">
            <canvas ref={trendRef}></canvas>
          </div>
        </div>
      </div>

      {/* Row 2: ABC + Insights */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* ABC Card */}
        <div style={cardStyle('#ea580c')} className="flex flex-col">
          <h3 className="text-xs font-extrabold text-violet-700 uppercase tracking-wider mb-4">📊 Clasificación ABC — % Movimientos</h3>
          <div className="relative h-48 mb-4">
            <canvas ref={abcRef}></canvas>
          </div>
          <div className="grid grid-cols-4 gap-2">
            {abcAll.map(r => {
              const cl = { 'A': 'text-emerald-600', 'B': 'text-amber-500', 'C': 'text-rose-500', 'X': 'text-slate-500' }[r.cls] || 'text-violet-600';
              const bg = { 'A': 'bg-emerald-50', 'B': 'bg-amber-50', 'C': 'bg-rose-50', 'X': 'bg-slate-50' }[r.cls] || 'bg-violet-50';
              const bd = { 'A': 'border-emerald-100', 'B': 'border-amber-100', 'C': 'border-rose-100', 'X': 'border-slate-100' }[r.cls] || 'border-violet-100';
              const montoFmt = r.monto > 0 ? (r.monto >= 1e6 ? `$${(r.monto / 1e6).toFixed(1)}M` : `$${(r.monto / 1e3).toFixed(0)}K`) : '—';
              return (
                <div key={r.cls} className={`text-center p-2 rounded-xl border ${bg} ${bd}`}>
                  <div className={`text-lg font-black ${cl}`}>{r.cls}</div>
                  <div className="text-[10px] text-slate-500 font-bold">{r.skus} SKUs</div>
                  <div className="text-[9px] text-blue-600 mt-0.5">{r.pctMovs}% movs</div>
                  <div className="text-[9px] text-amber-700 mt-0.5">{montoFmt}</div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Insights Card */}
        <div style={cardStyle('#6d28d9')} className="flex flex-col bg-white">
          <h3 className="text-xs font-extrabold text-violet-700 uppercase tracking-wider mb-4">💡 Insights & Acciones Recomendadas</h3>
          
          <div className="space-y-3 flex-1 overflow-y-auto max-h-72">
            {/* Veracidad Insight */}
            <div className={`flex items-start gap-3 p-3 rounded-xl border ${vPct >= 99 ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : vPct >= 95 ? 'bg-amber-50 border-amber-100 text-amber-800' : 'bg-red-50 border-red-100 text-red-800'}`}>
              <span className="text-base">{vPct >= 99 ? '✅' : vPct >= 95 ? '⚠️' : '🚨'}</span>
              <div className="text-xs">
                <span className="font-bold">Veracidad de inventario al {vPct}%</span> — {vPct >= 99 ? 'Excelente control y confiabilidad de los datos en SAP.' : vPct >= 95 ? 'Nivel aceptable. Se sugiere revisar discrepancias en conteos cíclicos.' : 'Discrepancia crítica. Acción correctiva urgente requerida.'}
              </div>
            </div>

            {/* MRP Insight */}
            <div className={`flex items-start gap-3 p-3 rounded-xl border ${nCortosCriticos === 0 && nCortos === 0 ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-red-50 border-red-100 text-red-800'}`}>
              <span className="text-base">{nCortosCriticos === 0 ? '🟢' : '🔴'}</span>
              <div className="text-xs">
                <span className="font-bold">Estatus de Cortos MRP:</span> {nCortosCriticos === 0 && nCortos === 0 ? 'Plan sin cortos. Cobertura asegurada para el ciclo actual.' : `${nCortosCriticos} cortos críticos sin stock en zona. Requiere compras de emergencia o activación de GLM.`}
              </div>
            </div>

            {/* Capacidad Insight */}
            <div className={`flex items-start gap-3 p-3 rounded-xl border ${parseFloat(capPct) >= 100 ? 'bg-red-50 border-red-100 text-red-800' : parseFloat(capPct) >= 85 ? 'bg-amber-50 border-amber-100 text-amber-800' : 'bg-emerald-50 border-emerald-100 text-emerald-800'}`}>
              <span className="text-base">🏭</span>
              <div className="text-xs">
                <span className="font-bold">Ocupación Física de Almacenes al {capPct}%:</span> {parseFloat(capPct) >= 100 ? 'Capacidad rebasada. Urge liberar posiciones físicas o trasladar material.' : parseFloat(capPct) >= 85 ? 'Espacio cercano al límite. Planificar reubicación.' : 'Disponibilidad de espacio adecuada en la red de almacenes.'}
              </div>
            </div>
          </div>

          <div className="mt-4 p-2 text-center bg-slate-50 border border-slate-100 rounded-xl text-[10px] text-slate-400 font-bold">
            Generado: {genDate} · Valor total de red: {fmxn(costInv)}
          </div>
        </div>
      </div>
    </div>
  );
}
