/**
 * dashboard.js
 * Handles the staff dashboard: stat card flashes,
 * release bar chart (Released / Returned), device‑monitoring bar chart,
 * and the WebSocket / AJAX‑poll real‑time connection.
 *
 * Expects a global DASHBOARD_INIT object injected by the Django template:
 *   const DASHBOARD_INIT = {
 *     released: <int>,
 *     returned: <int>,
 *     bar: { offices, serviceable, nonService, sealed, missing, incomplete }
 *   };
 */

(function () {
  'use strict';

  /* ── Stat card flash ──────────────────────────────────────────────────── */
  function flashStat(el, newValue) {
    if (!el) return;
    if (el.textContent.trim() === String(newValue)) return;
    el.textContent = newValue;
    el.classList.remove('flash');
    void el.offsetWidth; // reflow
    el.classList.add('flash');
    setTimeout(() => el.classList.remove('flash'), 300);
  }

  /* ═══════════════════════════════════════════════════════════════════════
     RELEASE BAR CHART  (replaces the old pie chart)
  ═══════════════════════════════════════════════════════════════════════ */
  let releaseBarChart = null;

  function drawReleaseBar(released, returned) {
    const canvas = document.getElementById('releaseBarChart');
    if (!canvas) return;
    if (releaseBarChart) { releaseBarChart.destroy(); releaseBarChart = null; }

    const total = released + returned;

    releaseBarChart = new Chart(canvas.getContext('2d'), {
      type: 'bar',
      data: {
        labels: ['Released', 'Returned'],
        datasets: total === 0
          ? [{ label: 'No Data', data: [0, 0], backgroundColor: '#334155' }]
          : [{
              label: 'Devices',
              data: [released, returned],
              backgroundColor: ['#f59e0b', '#22c55e'],
              borderWidth: 0,
              borderRadius: 6,
              barPercentage: 0.6,
            }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => `  ${ctx.label}: ${ctx.raw} (${total ? Math.round(ctx.raw / total * 100) : 0}%)`
            },
            bodyColor: '#fff',
            titleColor: '#fff',
            backgroundColor: '#1e293b',
            borderColor: 'rgba(255,255,255,0.1)',
            borderWidth: 1,
          }
        },
        scales: {
          x: {
            ticks: { color: '#fff', font: { size: 13 } },
            grid: { display: false }
          },
          y: {
            beginAtZero: true,
            ticks: { color: '#fff', stepSize: 1, precision: 0 },
            grid: { color: 'rgba(255,255,255,0.08)' }
          }
        }
      }
    });

    // Update legend numbers (the <strong> elements)
    const lr = document.getElementById('legend-released');
    const lt = document.getElementById('legend-returned');
    if (lr) lr.textContent = released;
    if (lt) lt.textContent = returned;
  }

  /* ── Bar chart (Device Monitoring by College / Office) ───────────────── */
  let barChart = null;

  function drawBar(bar) {
    const canvas = document.getElementById('barChart');
    if (!canvas) return;
    if (barChart) { barChart.destroy(); barChart = null; }

    const hasData     = bar.offices.length > 0;
    const LABEL_COLOR = '#ffffff';
    const GRID_COLOR  = 'rgba(255,255,255,0.08)';
    canvas.style.height = hasData
      ? Math.max(200, bar.offices.length * 48) + 'px'
      : '120px';

    barChart = new Chart(canvas.getContext('2d'), {
      type: 'bar',
      data: {
        labels: hasData ? bar.offices : ['No data yet'],
        datasets: hasData ? [
          { label: 'Serviceable',     data: bar.serviceable, backgroundColor: '#22c55e', stack: 'a' },
          { label: 'Non-Serviceable', data: bar.nonService,  backgroundColor: '#ef4444', stack: 'a' },
          { label: 'Sealed',          data: bar.sealed,      backgroundColor: '#6366f1', stack: 'a' },
          { label: 'Missing',         data: bar.missing,     backgroundColor: '#f59e0b', stack: 'a' },
          { label: 'Incomplete',      data: bar.incomplete,  backgroundColor: '#94a3b8', stack: 'a' },
        ] : [{ label: 'No data', data: [0], backgroundColor: '#334155', stack: 'a' }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              color: LABEL_COLOR, padding: 14,
              font: { size: 12 }, boxWidth: 12
            }
          },
          tooltip: {
            bodyColor: '#fff', titleColor: '#fff',
            backgroundColor: '#1e293b',
            borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1
          }
        },
        scales: {
          x: {
            stacked: true,
            ticks: { color: LABEL_COLOR, font: { size: 12 } },
            grid:   { color: GRID_COLOR },
            border: { color: GRID_COLOR }
          },
          y: {
            stacked: true,
            ticks: { color: LABEL_COLOR, font: { size: 12 } },
            grid:   { color: GRID_COLOR },
            border: { color: GRID_COLOR }
          }
        }
      }
    });
  }

  /* ── WebSocket message handler ────────────────────────────────────────── */
  function handleDashboardMessage(data) {
    if (data.type !== 'dashboard.update') return;

    flashStat(document.getElementById('stat-items'),   data.items_count);
    flashStat(document.getElementById('stat-borrows'),  data.active_borrows);
    flashStat(document.getElementById('stat-returns'),  data.total_returns);
    flashStat(document.getElementById('stat-pending'),  data.pending_count);

    window.dispatchEvent(new CustomEvent('invsys:pending_count',      { detail: data.pending_count }));
    window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));

    // Update release bar chart (Released / Returned)
    const released = typeof data.dm_released === 'number' ? data.dm_released : 0;
    const returned = typeof data.dm_returned  === 'number' ? data.dm_returned  : 0;
    drawReleaseBar(released, returned);

    if (data.bar) drawBar(data.bar);
  }

  /* ── Boot ─────────────────────────────────────────────────────────────── */
  function boot() {
    if (typeof DASHBOARD_INIT !== 'undefined') {
      drawReleaseBar(DASHBOARD_INIT.released, DASHBOARD_INIT.returned);
      drawBar(DASHBOARD_INIT.bar);
    }
    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined') {
      InvSysRT.connect('/ws/dashboard/', handleDashboardMessage, indicator);
    }
  }

  function loadChartJs(cb) {
    if (window.Chart) { cb(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js';
    s.onload = cb;
    document.head.appendChild(s);
  }

  window.addEventListener('pageshow', () => loadChartJs(boot));
})();