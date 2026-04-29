/* ── Toast Notifications ── */
function showToast(message, type = 'success') {
  const container = document.getElementById('toast-container') || createToastContainer();
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span class="toast-dot"></span><span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity   = '0';
    toast.style.transform = 'translateX(16px)';
    toast.style.transition = '0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

function createToastContainer() {
  const el = document.createElement('div');
  el.id = 'toast-container';
  document.body.appendChild(el);
  return el;
}

/* Auto-dismiss Django messages as toasts */
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.django-message').forEach(el => {
    const type = el.dataset.type === 'error' ? 'error' : 'success';
    showToast(el.textContent.trim(), type);
    el.remove();
  });
});

/* ── Table Search ── */
function initTableSearch(inputId, tableId) {
  const input = document.getElementById(inputId);
  const tbody = document.querySelector(`#${tableId} tbody`);
  if (!input || !tbody) return;
  input.addEventListener('input', () => {
    const q = input.value.toLowerCase();
    tbody.querySelectorAll('tr').forEach(row => {
      row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
  });
}

document.addEventListener('DOMContentLoaded', () => {
  initTableSearch('item-search', 'items-table');
  initTableSearch('transaction-search', 'transactions-table');
});

/* ── Quantity Bar Renderer ── */
function renderQtyBars() {
  document.querySelectorAll('[data-qty]').forEach(el => {
    const available = parseInt(el.dataset.qty, 10);
    const total     = parseInt(el.dataset.total, 10);
    if (isNaN(available) || isNaN(total) || total === 0) return;
    const pct = Math.round((available / total) * 100);
    const cls = pct <= 20 ? 'low' : pct <= 50 ? 'mid' : '';
    el.innerHTML = `
      <div class="qty-bar">
        <div class="qty-track">
          <div class="qty-fill ${cls}" style="width:${pct}%"></div>
        </div>
        <span>${available}/${total}</span>
      </div>`;
  });
}
document.addEventListener('DOMContentLoaded', renderQtyBars);

/* ── Borrow Form Live Check ── */
document.addEventListener('DOMContentLoaded', () => {
  const itemSelect = document.getElementById('id_item');
  const qtyInput   = document.getElementById('id_quantity_borrowed');
  const hintEl     = document.getElementById('qty-hint');
  if (!itemSelect || !qtyInput || !hintEl) return;

  const availMap = {};
  itemSelect.querySelectorAll('option[data-available]').forEach(opt => {
    availMap[opt.value] = parseInt(opt.dataset.available, 10);
  });

  function checkQty() {
    const available = availMap[itemSelect.value];
    if (available == null) return;
    const requested = parseInt(qtyInput.value, 10);
    hintEl.textContent = `Available: ${available} units`;
    hintEl.className   = 'form-hint';
    if (!isNaN(requested)) {
      if (requested > available) {
        hintEl.textContent = `⚠ Only ${available} units available`;
        hintEl.className   = 'form-error';
      } else if (requested < 1) {
        hintEl.textContent = 'Enter at least 1';
        hintEl.className   = 'form-error';
      }
    }
  }

  itemSelect.addEventListener('change', checkQty);
  qtyInput.addEventListener('input', checkQty);
  checkQty();
});

/* ── Badge Live Updates (from WebSocket / AJAX) ── */
(function() {
  // Pending requests badge
  const pendingBadge = document.getElementById('nav-pending-badge');
  if (pendingBadge) {
    window.addEventListener('invsys:pending_count', (e) => {
      const count = e.detail;
      // GUARD: skip if not a real number — undefined/null/NaN must never hide the badge
      if (typeof count !== 'number' || isNaN(count)) return;
      pendingBadge.textContent   = count > 0 ? count : '';
      pendingBadge.style.display = count > 0 ? '' : 'none';
    });
  }

  // Graduation warning badge
  const gradBadge = document.getElementById('nav-grad-badge');
  if (gradBadge) {
    window.addEventListener('invsys:grad_warning_count', (e) => {
      const count = e.detail;
      // GUARD: skip if not a real number — undefined/null/NaN must never hide the badge
      if (typeof count !== 'number' || isNaN(count)) return;
      gradBadge.textContent      = count >= 0 ? String(count) : '0';
      gradBadge.style.display    = count > 0 ? '' : 'none';
      if (count > 0) {
        gradBadge.classList.add('grad-badge-pulse');
      } else {
        gradBadge.classList.remove('grad-badge-pulse');
      }
    });
  }
})();