/**
 * borrow_requests.js
 * Handles real-time WebSocket updates for the Borrow Requests page
 * and the client-side search filter.
 */

(function () {
  'use strict';

  /* ── Helpers ──────────────────────────────────────────────────────────── */
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>]/g, m =>
      m === '&' ? '&amp;' : m === '<' ? '&lt;' : '&gt;'
    );
  }

  function showToast(message, type) {
    let toast = document.querySelector('.custom-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.className = 'custom-toast';
      toast.style.cssText = `
        position:fixed;bottom:20px;right:20px;
        padding:12px 20px;border-radius:8px;font-weight:500;
        z-index:2000;opacity:0;transform:translateY(20px);
        transition:all 0.3s ease;font-family:monospace;font-size:13px;
        pointer-events:none;`;
      document.body.appendChild(toast);
    }
    toast.style.background = type === 'error' ? '#ff4444' : '#00e5a0';
    toast.style.color       = type === 'error' ? '#fff'    : '#000';
    toast.textContent = message;
    toast.style.opacity   = '1';
    toast.style.transform = 'translateY(0)';
    setTimeout(() => {
      toast.style.opacity   = '0';
      toast.style.transform = 'translateY(20px)';
    }, 3000);
  }

  function formatPhilippineTime(dateString) {
    if (!dateString || dateString === '—') return '—';
    try {
      const date = new Date(dateString);
      if (!isNaN(date.getTime())) {
        return date.toLocaleString('en-US', {
          year: 'numeric', month: 'short', day: 'numeric',
          hour: 'numeric', minute: '2-digit', hour12: true,
          timeZone: 'Asia/Manila'
        });
      }
    } catch (e) {}
    return dateString;
  }

  /* ── Search filter ────────────────────────────────────────────────────── */
  let searchQuery = '';

  function applySearchFilter() {
    const rows = document.querySelectorAll('#requests-tbody tr[id^="req-row-"]');
    rows.forEach(row => {
      row.style.display = (!searchQuery || row.textContent.toLowerCase().includes(searchQuery)) ? '' : 'none';
    });
  }

  /* ── WebSocket real-time updates ──────────────────────────────────────── */
  const CSRF = document.cookie.match(/csrftoken=([^;]+)/)?.[1] ?? '';

  const knownIds = new Set(
    [...document.querySelectorAll('#requests-tbody tr[id^="req-row-"]')]
      .map(r => r.id.replace('req-row-', ''))
  );

  function renderRequests(requests) {
    const tbody = document.getElementById('requests-tbody');
    if (!tbody) return;

    const badge = document.getElementById('pending-count-badge');

    // Remove rows no longer pending
    const incomingIds = new Set(requests.map(r => String(r.id)));
    tbody.querySelectorAll('tr[id^="req-row-"]').forEach(tr => {
      if (!incomingIds.has(tr.id.replace('req-row-', ''))) tr.remove();
    });

    requests.forEach(req => {
      const existing = document.getElementById('req-row-' + req.id);
      const isNew    = !knownIds.has(String(req.id));
      knownIds.add(String(req.id));

      if (!existing) {
        const tr = document.createElement('tr');
        tr.id = 'req-row-' + req.id;
        tr.setAttribute('data-tx-id',   (req.transaction_id || '').toLowerCase());
        tr.setAttribute('data-borrower', (req.borrower_name || '').toLowerCase());
        tr.setAttribute('data-office',   (req.office_college || '').toLowerCase());
        tr.setAttribute('data-item',     (req.item_name || '').toLowerCase());

        tr.innerHTML = `
          <td style="text-align:center"><span class="badge badge-blue">#${escapeHtml(req.transaction_id)}</span></td>
          <td style="text-align:center;font-weight:600">${escapeHtml(req.borrower_name)}</td>
          <td style="text-align:center">${escapeHtml(req.office_college)}</td>
          <td style="text-align:center">${escapeHtml(req.item_name || '—')}</td>
          <td style="text-align:center">${req.quantity}</td>
          <td style="text-align:center;color:var(--muted)">${formatPhilippineTime(req.created_at)}</td>
          <td style="text-align:center">
            <div style="display:flex;gap:6px;justify-content:center">
              <a href="/borrow/confirm/${req.id}/" class="btn btn-primary btn-sm">Accept</a>
              <form method="post" action="/requests/${req.id}/decline/" style="display:inline">
                <input type="hidden" name="csrfmiddlewaretoken" value="${CSRF}"/>
                <button type="submit" class="btn btn-danger btn-sm">Decline</button>
              </form>
            </div>
          </td>
        `;

        const emptyRow = document.getElementById('req-empty-row');
        if (emptyRow) emptyRow.remove();
        tbody.prepend(tr);

        if (isNew) {
          tr.classList.add('row-new');
          setTimeout(() => tr.classList.remove('row-new'), 1600);
          showToast(`New request from ${req.borrower_name}`, 'success');
        }
      }
    });

    // Show empty state if needed
    if (tbody.querySelectorAll('tr[id^="req-row-"]').length === 0 && !document.getElementById('req-empty-row')) {
      const tr = document.createElement('tr');
      tr.id = 'req-empty-row';
      tr.innerHTML = '<td colspan="7"><div class="empty-state" style="padding:40px;text-align:center"><p>No pending borrow requests.</p></div></td>';
      tbody.appendChild(tr);
    }

    const count = requests.length;
    if (badge) {
      badge.textContent   = count;
      badge.style.display = count > 0 ? '' : 'none';
    }

    applySearchFilter();
  }

  function handleMessage(data) {
    if (data.type !== 'borrow_requests.update') return;
    renderRequests(data.pending);
    window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.count }));
    window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));
  }

  /* ── DOMContentLoaded ─────────────────────────────────────────────────── */
  document.addEventListener('DOMContentLoaded', () => {
    const searchInput = document.getElementById('request-search');
    if (searchInput) {
      searchInput.addEventListener('input', e => {
        searchQuery = e.target.value.toLowerCase().trim();
        applySearchFilter();
      });
    }

    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined') {
      InvSysRT.connect('/ws/borrow-requests/', handleMessage, indicator);
    }
  });
})();