/**
 * dashboard.js
 * Handles the staff dashboard: stat card flashes, pie chart, bar chart,
 * and the WebSocket / AJAX-poll real-time connection.
 *
 * Expects a global DASHBOARD_INIT object injected by the Django template:
 *   const DASHBOARD_INIT = {
 *     available: <int>,
 *     borrowed:  <int>,
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

  /* ── Pie chart ────────────────────────────────────────────────────────── */
  let pieChart = null;

  function drawPie(available, borrowed) {
    const canvas = document.getElementById('pieChart');
    if (!canvas) return;
    if (pieChart) { pieChart.destroy(); pieChart = null; }

    const total = available + borrowed;

    if (total === 0) {
      pieChart = new Chart(canvas.getContext('2d'), {
        type: 'pie',
        data: {
          labels: ['No Data'],
          datasets: [{ data: [1], backgroundColor: ['#334155'], borderWidth: 2, borderColor: '#1e293b' }]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: { label: () => ' No data' },
              bodyColor: '#fff', titleColor: '#fff',
              backgroundColor: '#1e293b', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1
            }
          }
        }
      });
    } else {
      pieChart = new Chart(canvas.getContext('2d'), {
        type: 'pie',
        data: {
          labels: ['Available', 'Borrowed'],
          datasets: [{
            data: [available, borrowed],
            backgroundColor: ['#22c55e', '#f59e0b'],
            borderWidth: 2, borderColor: '#1e293b'
          }]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: { label: ctx => `  ${ctx.label}: ${ctx.parsed} of ${total}` },
              bodyColor: '#fff', titleColor: '#fff',
              backgroundColor: '#1e293b', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1
            }
          }
        }
      });
    }

    const la = document.getElementById('legend-available');
    const lb = document.getElementById('legend-borrowed');
    if (la) la.textContent = available;
    if (lb) lb.textContent = borrowed;
  }

  /* ── Bar chart ────────────────────────────────────────────────────────── */
  let barChart = null;

  function drawBar(bar) {
    const canvas = document.getElementById('barChart');
    if (!canvas) return;
    if (barChart) { barChart.destroy(); barChart = null; }

    const hasData     = bar.offices.length > 0;
    const LABEL_COLOR = '#ffffff';
    const GRID_COLOR  = 'rgba(255,255,255,0.08)';
    canvas.style.height = hasData ? Math.max(200, bar.offices.length * 48) + 'px' : '120px';

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
            labels: { color: LABEL_COLOR, padding: 14, font: { size: 12 }, boxWidth: 12 }
          },
          tooltip: {
            bodyColor: '#fff', titleColor: '#fff',
            backgroundColor: '#1e293b', borderColor: 'rgba(255,255,255,0.1)', borderWidth: 1
          }
        },
        scales: {
          x: { stacked: true, ticks: { color: LABEL_COLOR, font: { size: 12 } }, grid: { color: GRID_COLOR }, border: { color: GRID_COLOR } },
          y: { stacked: true, ticks: { color: LABEL_COLOR, font: { size: 12 } }, grid: { color: GRID_COLOR }, border: { color: GRID_COLOR } }
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

    window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.pending_count }));

    drawPie(data.available_qty, data.borrowed_qty);
    if (data.bar) drawBar(data.bar);
  }

  /* ── Boot ─────────────────────────────────────────────────────────────── */
  function boot() {
    if (typeof DASHBOARD_INIT !== 'undefined') {
      drawPie(DASHBOARD_INIT.available, DASHBOARD_INIT.borrowed);
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