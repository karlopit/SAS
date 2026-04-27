/**
 * borrow_management.js
 * Handles:
 *  - Return modal (open, close, device list, confirm)
 *  - Filters (college, officer, borrower name/type, status, search)
 *  - Real-time WebSocket updates for items table and transactions table
 */

(function () {
  'use strict';

  /* ── Helpers ──────────────────────────────────────────────────────────── */
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>]/g, function (m) {
      return m === '&' ? '&amp;' : m === '<' ? '&lt;' : '&gt;';
    });
  }

  function showToast(message, type) {
    let toast = document.querySelector('.custom-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.className = 'custom-toast';
      document.body.appendChild(toast);
    }
    toast.style.background = type === 'error' ? '#ff4444' : '#00e5a0';
    toast.style.color       = type === 'error' ? '#fff'    : '#000';
    toast.textContent = message;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
  }

  function getCsrf() {
    return document.cookie.match(/csrftoken=([^;]+)/)?.[1] ?? '';
  }

  /* ══════════════════════════════════════════════════════════════════════
     RETURN MODAL
  ══════════════════════════════════════════════════════════════════════ */
  let _returnTxId    = null;
  let _returnDevices = [];

  window.openReturnModal = function (btn) {
    _returnTxId = btn.dataset.txId;
    const borrower = btn.dataset.borrower;
    const qty      = btn.dataset.qty;

    document.getElementById('return-modal-subtitle').textContent =
      `${borrower} · ${qty} device(s) borrowed`;

    document.getElementById('return-modal-loading').style.display  = 'block';
    document.getElementById('return-device-list').style.display    = 'none';
    document.getElementById('return-empty-state').style.display    = 'none';
    document.getElementById('return-device-rows').innerHTML        = '';
    document.getElementById('return-footer-note').textContent      = '';

    const confirmBtn = document.getElementById('confirm-return-btn');
    confirmBtn.disabled  = true;
    confirmBtn.innerHTML = `<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:14px;height:14px">
      <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/>
    </svg> Confirm Return`;

    document.getElementById('returnModal').style.display = 'flex';

    fetch(`/transaction/${_returnTxId}/devices/`, { credentials: 'same-origin' })
      .then(r => r.json())
      .then(data => {
        _returnDevices = data.devices || [];
        renderReturnDevices(_returnDevices);
      })
      .catch(() => {
        document.getElementById('return-modal-loading').style.display = 'none';
        document.getElementById('return-footer-note').textContent = 'Error loading devices.';
        showToast('Error loading devices', 'error');
      });
  };

  window.closeReturnModal = function (e) {
    if (e && e.target !== document.getElementById('returnModal')) return;
    document.getElementById('returnModal').style.display = 'none';
    _returnTxId    = null;
    _returnDevices = [];
  };

  function renderReturnDevices(devices) {
    document.getElementById('return-modal-loading').style.display = 'none';

    const allReturned = devices.every(d => d.returned);
    if (devices.length === 0 || allReturned) {
      document.getElementById('return-empty-state').style.display = 'block';
      return;
    }

    document.getElementById('return-device-list').style.display = 'block';

    const container     = document.getElementById('return-device-rows');
    container.innerHTML = '';
    const totalReturned = devices.filter(d => d.returned).length;

    devices.forEach((d) => {
      const row = document.createElement('div');
      row.className = 'return-device-row' + (d.returned ? ' already-returned' : '');

      if (!d.returned) {
        row.addEventListener('click', (e) => {
          if (e.target.tagName === 'INPUT') return;
          const cb = row.querySelector('input[type="checkbox"]');
          cb.checked = !cb.checked;
          onDeviceCheckChange(row, cb.checked);
        });
      }

      const statusHtml = d.returned
        ? `<span style="font-size:11px;color:#00e5a0;font-weight:600">✓ Returned</span>`
        : `<span style="font-size:11px;color:var(--muted)">Pending</span>`;

      row.innerHTML = `
        <label class="return-device-label" onclick="event.stopPropagation()">
          <input type="checkbox"
                 data-device-id="${escapeHtml(d.id ?? '')}"
                 data-serial="${escapeHtml(d.serial_number)}"
                 ${d.returned ? 'checked disabled' : ''}
                 onchange="onDeviceCheckChange(this.closest('.return-device-row'), this.checked)"/>
        </label>
        <span class="return-device-serial">${escapeHtml(d.serial_number)}</span>
        <span class="return-device-box">${escapeHtml(d.box_number || '—')}</span>
        <span class="return-device-status">${statusHtml}</span>
      `;
      container.appendChild(row);
    });

    updateReturnProgress(totalReturned, devices.length);
    updateReturnConfirmBtn();
  }

  window.onDeviceCheckChange = function (row, checked) {
    if (checked) row.classList.add('selected');
    else {
      row.classList.remove('selected');
      document.getElementById('select-all-devices').checked = false;
    }
    updateReturnConfirmBtn();
  };

  window.toggleSelectAll = function (masterCb) {
    document.querySelectorAll('#return-device-rows input[type="checkbox"]:not(:disabled)')
      .forEach(cb => {
        cb.checked = masterCb.checked;
        const row  = cb.closest('.return-device-row');
        if (masterCb.checked) row.classList.add('selected');
        else                  row.classList.remove('selected');
      });
    updateReturnConfirmBtn();
  };

  function updateReturnConfirmBtn() {
    const checkedCount = document.querySelectorAll(
      '#return-device-rows input[type="checkbox"]:not(:disabled):checked'
    ).length;
    document.getElementById('confirm-return-btn').disabled = checkedCount === 0;
    document.getElementById('return-selected-count').textContent = `${checkedCount} selected`;
  }

  function updateReturnProgress(returned, total) {
    const pct = total > 0 ? Math.round((returned / total) * 100) : 0;
    document.getElementById('return-progress-bar').style.width = pct + '%';
    document.getElementById('return-progress-label').textContent = `${returned} / ${total} returned`;
  }

  window.confirmReturn = async function () {
    const checkedBoxes = document.querySelectorAll(
      '#return-device-rows input[type="checkbox"]:not(:disabled):checked'
    );
    if (checkedBoxes.length === 0) return;

    const serials   = [];
    const deviceIds = [];
    checkedBoxes.forEach(cb => {
      serials.push(cb.dataset.serial);
      if (cb.dataset.deviceId && cb.dataset.deviceId !== 'null' && cb.dataset.deviceId !== 'undefined') {
        deviceIds.push(parseInt(cb.dataset.deviceId));
      }
    });

    const btn = document.getElementById('confirm-return-btn');
    btn.disabled  = true;
    btn.innerHTML = 'Processing…';

    try {
      const payload = { serials };
      if (deviceIds.length > 0) payload.device_ids = deviceIds;

      const resp = await fetch(`/transaction/${_returnTxId}/return-devices/`, {
        method:      'POST',
        headers:     { 'Content-Type': 'application/json', 'X-CSRFToken': getCsrf() },
        body:        JSON.stringify(payload),
        credentials: 'same-origin',
      });
      const data = await resp.json();
      if (!data.ok) throw new Error(data.error || 'Server error');

      // Update table row in place
      const txRow = document.getElementById('tx-row-' + _returnTxId);
      if (txRow) {
        const qtyDisplay = txRow.querySelector('.returned-qty-display');
        if (qtyDisplay) qtyDisplay.textContent = data.returned_qty;

        const statusCell = txRow.querySelector('.status-cell');
        if (statusCell) {
          if (data.fully_returned) {
            statusCell.innerHTML = '<span class="badge badge-returned-full">✓ Returned</span>';
          } else {
            const borrowerName = txRow.querySelector('td:nth-child(2)')?.textContent?.trim() || '';
            const qtyBorrowed  = txRow.querySelector('td:nth-child(7)')?.textContent?.trim() || '';
            statusCell.innerHTML = `
              <button class="btn-return-action"
                      data-tx-id="${_returnTxId}"
                      data-borrower="${escapeHtml(borrowerName)}"
                      data-qty="${qtyBorrowed}"
                      onclick="openReturnModal(this)">↩ Return</button>`;
          }
        }

        const returnedOnCell = txRow.querySelector('.returned-on-cell');
        if (returnedOnCell && data.returned_at && data.returned_at !== '—') {
          returnedOnCell.textContent = data.returned_at;
        }
        txRow.dataset.status = data.fully_returned ? 'returned' : 'borrowed';
      }

      showToast(`${serials.length} device(s) returned successfully`, 'success');
      document.getElementById('returnModal').style.display = 'none';

    } catch (err) {
      showToast('Error: ' + err.message, 'error');
    } finally {
      btn.disabled  = false;
      btn.innerHTML = `<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:14px;height:14px">
        <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/>
      </svg> Confirm Return`;
    }
  };

  /* ══════════════════════════════════════════════════════════════════════
     FILTERS
  ══════════════════════════════════════════════════════════════════════ */
  let searchQuery = '', filterCollege = '', filterOfficer = '',
      filterBorrowerName = '', filterBorrowerType = '', filterStatus = '';

  function applyFilters() {
    const rows  = document.querySelectorAll('#transactions-tbody tr[id^="tx-row-"]');
    let visible = 0;
    const total = rows.length;

    rows.forEach(row => {
      const text         = row.textContent.toLowerCase();
      const college      = (row.dataset.college || '').toLowerCase();
      const officer      = (row.dataset.officer || '').toLowerCase();
      const borrowerName = (row.dataset.borrowerName || '').toLowerCase();
      const borrowerType = (row.dataset.borrowerType || '').toLowerCase();
      const status       = (row.dataset.status || '').toLowerCase();

      const show = (
        (!searchQuery        || text.includes(searchQuery)) &&
        (!filterCollege      || college      === filterCollege.toLowerCase()) &&
        (!filterOfficer      || officer      === filterOfficer.toLowerCase()) &&
        (!filterBorrowerName || borrowerName === filterBorrowerName.toLowerCase()) &&
        (!filterBorrowerType || borrowerType === filterBorrowerType.toLowerCase()) &&
        (!filterStatus       || status       === filterStatus.toLowerCase())
      );
      row.style.display = show ? '' : 'none';
      if (show) visible++;
    });

    const hasFilter = searchQuery || filterCollege || filterOfficer ||
                      filterBorrowerName || filterBorrowerType || filterStatus;
    const bar = document.getElementById('filter-status-bar');
    if (bar) bar.style.display = hasFilter ? '' : 'none';
    const fc = document.getElementById('filter-count');
    const ft = document.getElementById('filter-total');
    if (fc) fc.textContent = visible;
    if (ft) ft.textContent = total;
  }

  function dedupeSelect(sel) {
    if (!sel) return;
    const seen = new Set([sel.options[0]?.value ?? '']);
    for (let i = sel.options.length - 1; i > 0; i--) {
      if (seen.has(sel.options[i].value)) sel.remove(i);
      else seen.add(sel.options[i].value);
    }
  }

  window.clearFilters = function () {
    searchQuery = filterCollege = filterOfficer =
    filterBorrowerName = filterBorrowerType = filterStatus = '';
    const si = document.getElementById('transaction-search');
    if (si) si.value = '';
    ['filter-college','filter-officer','filter-borrower-name',
     'filter-borrower-type','filter-status'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.selectedIndex = 0;
    });
    applyFilters();
  };

  /* ══════════════════════════════════════════════════════════════════════
     REAL-TIME RENDERING
  ══════════════════════════════════════════════════════════════════════ */
  function renderItems(items) {
    const tbody = document.getElementById('items-tbody');
    if (!tbody) return;
    tbody.innerHTML = items.map((item, i) => {
      const pct   = item.quantity ? Math.round((item.available_quantity / item.quantity) * 100) : 0;
      const cls   = pct <= 20 ? 'low' : pct <= 50 ? 'mid' : '';
      const badge = item.available_quantity === 0
        ? '<span class="badge badge-red">Out of Stock</span>'
        : '<span class="badge badge-green">In Stock</span>';
      return `<tr>
        <td style="text-align:center;color:var(--muted)">${i + 1}</td>
        <td style="text-align:center;font-weight:600">${escapeHtml(item.name)}</td>
        <td style="text-align:center;color:var(--muted);font-size:12px">${escapeHtml(item.serial || '—').replace(/\n/g,'<br>')}</td>
        <td style="text-align:center;color:var(--muted)">${escapeHtml((item.description || '—').substring(0,40))}</td>
        <td style="text-align:center">
          <div class="qty-bar">
            <div class="qty-track"><div class="qty-fill ${cls}" style="width:${pct}%"></div></div>
            <span>${item.available_quantity}/${item.quantity}</span>
          </div>
        </td>
        <td style="text-align:center">${badge}</td>
      </tr>`;
    }).join('');
  }

  const knownTxIds = new Set(
    [...document.querySelectorAll('#transactions-tbody tr[id^="tx-row-"]')]
      .map(r => r.id.replace('tx-row-', ''))
  );

  function borrowerTypeBadge(type) {
    if (type === 'student')  return '<span class="badge badge-green">Student</span>';
    if (type === 'employee') return '<span class="badge badge-purple">Employee</span>';
    return '<span style="color:var(--muted)">—</span>';
  }

  function statusCellHtml(tx) {
    if (tx.fully_returned) {
      return '<span class="badge badge-returned-full">✓ Returned</span>';
    }
    return `<button class="btn-return-action"
                    data-tx-id="${tx.id}"
                    data-borrower="${escapeHtml(tx.borrower_name)}"
                    data-qty="${tx.qty_borrowed}"
                    onclick="openReturnModal(this)">↩ Return</button>`;
  }

  function renderTransactions(txs) {
    const tbody = document.getElementById('transactions-tbody');
    if (!tbody) return;

    txs.forEach(tx => {
      const existing = document.getElementById('tx-row-' + tx.id);
      const isNew    = !knownTxIds.has(String(tx.id));
      knownTxIds.add(String(tx.id));

      if (existing) {
        if (_returnTxId === String(tx.id)) return;

        existing.dataset.college      = tx.office_college;
        existing.dataset.officer      = tx.accountable_officer;
        existing.dataset.borrowerName = tx.borrower_name;
        existing.dataset.borrowerType = tx.borrower_type || '';
        existing.dataset.status       = tx.fully_returned ? 'returned' : 'borrowed';

        const qtyDisplay = existing.querySelector('.returned-qty-display');
        if (qtyDisplay) qtyDisplay.textContent = tx.returned_qty;

        const onCell = existing.querySelector('.returned-on-cell');
        if (onCell) onCell.textContent = tx.returned_at || '—';

        const statusCell = existing.querySelector('.status-cell');
        if (statusCell) statusCell.innerHTML = statusCellHtml(tx);

      } else {
        const tr = document.createElement('tr');
        tr.id = 'tx-row-' + tx.id;
        tr.dataset.college      = tx.office_college;
        tr.dataset.officer      = tx.accountable_officer;
        tr.dataset.borrowerName = tx.borrower_name;
        tr.dataset.borrowerType = tx.borrower_type || '';
        tr.dataset.status       = tx.fully_returned ? 'returned' : 'borrowed';

        tr.innerHTML = `
          <td style="text-align:center"><span class="badge badge-blue">${escapeHtml(tx.tx_id)}</span></td>
          <td style="text-align:center;font-weight:600">${escapeHtml(tx.borrower_name)}</td>
          <td style="text-align:center">${borrowerTypeBadge(tx.borrower_type)}</td>
          <td style="text-align:center;font-weight:600;color:var(--accent2)">${escapeHtml(tx.accountable_officer)}</td>
          <td style="text-align:center">${escapeHtml(tx.office_college)}</td>
          <td style="text-align:center">${escapeHtml(tx.item_name)}</td>
          <td style="text-align:center">${tx.qty_borrowed}</td>
          <td style="text-align:center">
            <span class="returned-qty-display">${tx.returned_qty}</span>
            <span style="color:var(--muted);font-size:11px"> / ${tx.qty_borrowed}</span>
          </td>
          <td style="text-align:center;color:var(--muted)">${escapeHtml(tx.borrowed_at)}</td>
          <td class="returned-on-cell" style="text-align:center;color:var(--muted);font-size:12px">${escapeHtml(tx.returned_at || '—')}</td>
          <td class="status-cell" style="text-align:center">${statusCellHtml(tx)}</td>
        `;

        const emptyRow = document.getElementById('tx-empty-row');
        if (emptyRow) emptyRow.remove();
        tbody.prepend(tr);

        if (isNew) {
          tr.classList.add('row-new');
          setTimeout(() => tr.classList.remove('row-new'), 1600);
        }
      }
    });

    applyFilters();
  }

  function handleMessage(data) {
  if (data.type !== 'borrow_management.update') return;
  renderItems(data.items);
  renderTransactions(data.transactions);
  window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.pending_count }));
  window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));
}

  /* ── DOMContentLoaded init ────────────────────────────────────────────── */
  document.addEventListener('DOMContentLoaded', () => {
    dedupeSelect(document.getElementById('filter-college'));
    dedupeSelect(document.getElementById('filter-officer'));
    dedupeSelect(document.getElementById('filter-borrower-name'));

    const bind = (id, setter) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', e => { setter(e.target.value); applyFilters(); });
    };
    const bindInput = (id, setter) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('input', e => { setter(e.target.value.toLowerCase().trim()); applyFilters(); });
    };

    bindInput('transaction-search', v => searchQuery = v);
    bind('filter-college',       v => filterCollege      = v);
    bind('filter-officer',       v => filterOfficer      = v);
    bind('filter-borrower-name', v => filterBorrowerName = v);
    bind('filter-borrower-type', v => filterBorrowerType = v);
    bind('filter-status',        v => filterStatus       = v);

    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined') {
      InvSysRT.connect('/ws/borrow-management/', handleMessage, indicator);
    }
  });
})();