/**
 * device_monitoring.js
 * - saveAllRows() now uses JSON (fixes TooManyFieldsSent)
 * - Scroll button is sticky + centered
 */

(function () {
  'use strict';

  /* ==================== TOAST ==================== */
  function showToast(message, type) {
    let el = document.getElementById('dm-toast');
    if (!el) {
      el = document.createElement('div');
      el.id = 'dm-toast';
      el.style.cssText = `
        position:fixed;bottom:24px;right:24px;z-index:9999;
        padding:12px 20px;border-radius:10px;font-family:var(--font-mono);
        font-size:13px;font-weight:600;pointer-events:none;
        opacity:0;transform:translateY(12px);
        transition:opacity .25s ease,transform .25s ease;
        box-shadow:0 8px 24px rgba(0,0,0,.4);`;
      document.body.appendChild(el);
    }
    el.style.background = type === 'error' ? 'rgba(255,76,76,.95)' : 'rgba(0,229,160,.95)';
    el.style.color      = type === 'error' ? '#fff' : '#000';
    el.textContent = message;
    el.style.opacity   = '1';
    el.style.transform = 'translateY(0)';
    clearTimeout(el._timer);
    el._timer = setTimeout(() => {
      el.style.opacity   = '0';
      el.style.transform = 'translateY(12px)';
    }, 3200);
  }

  /* ==================== CHECKBOX HELPERS ==================== */
  window.syncCheck = function (cb) {
    cb.previousElementSibling.value = cb.checked ? 'on' : 'off';
  };

  function getChecksInRow(row) {
    const fields = ['serviceable', 'non_serviceable', 'sealed', 'missing', 'incomplete'];
    const result = {};
    fields.forEach(f => {
      const hidden = row.querySelector(`input[type=hidden][name="${f}"]`);
      result[f] = { hidden, cb: hidden ? hidden.nextElementSibling : null };
    });
    return result;
  }

  window.handleDmCheck = function (cb, field) {
    const row = cb.closest('tr');
    const c   = getChecksInRow(row);
    if (cb.checked) {
      if (['non_serviceable', 'missing', 'incomplete'].includes(field)) {
        Object.keys(c).forEach(k => {
          if (k !== field) { c[k].cb.checked = false; c[k].hidden.value = 'off'; }
        });
      }
      if (field === 'serviceable' || field === 'sealed') {
        ['non_serviceable', 'missing', 'incomplete'].forEach(k => {
          c[k].cb.checked = false; c[k].hidden.value = 'off';
        });
      }
    }
    applyLockState(row);
    updateRowDataAttrsFromRow(row);
    applyDmFilters();
  };

  function applyLockState(row) {
    const c = getChecksInRow(row);
    if (!c.serviceable.cb) return;
    const exclusiveOn = c.non_serviceable.cb.checked || c.missing.cb.checked || c.incomplete.cb.checked;
    const safeOn      = c.serviceable.cb.checked || c.sealed.cb.checked;
    if (exclusiveOn) {
      Object.values(c).forEach(x => { x.cb.disabled = !x.cb.checked; });
    } else if (safeOn) {
      ['non_serviceable', 'missing', 'incomplete'].forEach(k => { c[k].cb.disabled = true; });
      c.serviceable.cb.disabled = false;
      c.sealed.cb.disabled      = false;
    } else {
      Object.values(c).forEach(x => { x.cb.disabled = false; });
    }
  }

  function updateRowDataAttrsFromRow(row) {
    if (!row) return;
    const c = getChecksInRow(row);
    row.dataset.serviceable    = c.serviceable.cb?.checked     ? '1' : '0';
    row.dataset.nonServiceable = c.non_serviceable.cb?.checked ? '1' : '0';
    row.dataset.sealed         = c.sealed.cb?.checked          ? '1' : '0';
    row.dataset.missing        = c.missing.cb?.checked         ? '1' : '0';
    row.dataset.incomplete     = c.incomplete.cb?.checked      ? '1' : '0';
  }

  /* ==================== BOX NUMBER SORT ==================== */
  function boxSortKey(row) {
    const raw = (row.dataset.box || row.querySelector('input[name="box_number"]')?.value || '').trim();
    if (!raw) return Infinity;
    const match = raw.match(/(\d+)/);
    return match ? parseInt(match[1], 10) : Infinity;
  }

  function sortTableByBoxNumber() {
    const tbody = document.getElementById('dm-tbody');
    if (!tbody) return;
    const rows = Array.from(tbody.querySelectorAll('tr[data-row-id]'));
    rows.sort((a, b) => boxSortKey(a) - boxSortKey(b));
    rows.forEach(tr => tbody.appendChild(tr));
  }

  /* ==================== FILTER DROPDOWNS ==================== */
  function toTitleCase(str) {
    if (!str) return '';
    return str.replace(/\w\S*/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
  }

  function populateFilterDropdowns() {
    const collegeSelect = document.getElementById('dm-filter-college');
    const officerSelect = document.getElementById('dm-filter-officer');
    const mrSelect      = document.getElementById('dm-filter-mr');
    const ptrSelect     = document.getElementById('dm-filter-ptr');
    if (!collegeSelect || !officerSelect || !mrSelect || !ptrSelect) return;

    const savedCollege = collegeSelect.value;
    const savedOfficer = officerSelect.value;
    const savedMr      = mrSelect.value;
    const savedPtr     = ptrSelect.value;

    while (collegeSelect.options.length > 1) collegeSelect.remove(1);
    while (officerSelect.options.length > 1) officerSelect.remove(1);
    while (mrSelect.options.length > 1)      mrSelect.remove(1);
    while (ptrSelect.options.length > 1)     ptrSelect.remove(1);

    const collegeMap = new Map();
    const officerMap = new Map();
    const mrMap      = new Map();
    const ptrMap     = new Map();

    document.querySelectorAll('#dm-tbody tr[data-row-id]').forEach(row => {
      const collegeRaw = row.dataset.collegeRaw || '';
      const officerRaw = row.dataset.officerRaw || '';
      const mrRaw      = row.dataset.mr          || '';
      const ptrRaw     = row.dataset.ptr         || '';
      if (collegeRaw) collegeMap.set(collegeRaw, collegeRaw);
      if (officerRaw) officerMap.set(officerRaw, officerRaw);
      if (mrRaw)      mrMap.set(mrRaw, toTitleCase(mrRaw));
      if (ptrRaw)     ptrMap.set(ptrRaw, toTitleCase(ptrRaw));
    });

    const addOptsFromMap = (sel, map, saved) => {
      [...map.keys()].sort((a, b) => a.localeCompare(b)).forEach(raw => {
        const o = document.createElement('option');
        o.value = raw; o.textContent = map.get(raw);
        if (raw === saved) o.selected = true;
        sel.appendChild(o);
      });
    };
    addOptsFromMap(collegeSelect, collegeMap, savedCollege);
    addOptsFromMap(officerSelect, officerMap, savedOfficer);
    addOptsFromMap(mrSelect,      mrMap,      savedMr);
    addOptsFromMap(ptrSelect,     ptrMap,     savedPtr);
  }

  /* ==================== SEARCH ==================== */
  function rowMatchesSearch(row, query) {
    if (!query) return true;
    const tds = row.querySelectorAll('td');
    let text = '';
    tds.forEach(td => { text += ' ' + td.textContent; });
    return text.toLowerCase().includes(query);
  }

  /* ==================== APPLY FILTERS ==================== */
  function applyDmFilters() {
    const search       = (document.getElementById('dm-search')?.value || '').toLowerCase().trim();
    const college      = document.getElementById('dm-filter-college')?.value || '';
    const borrowerType = document.getElementById('dm-filter-borrower-type')?.value || '';
    const officer      = document.getElementById('dm-filter-officer')?.value || '';
    const mr           = document.getElementById('dm-filter-mr')?.value || '';
    const ptr          = document.getElementById('dm-filter-ptr')?.value || '';
    const releaseF     = document.getElementById('dm-filter-release')?.value || '';
    const status       = document.getElementById('dm-filter-status')?.value || '';

    const rows = document.querySelectorAll('#dm-tbody tr[data-row-id]');
    let visible = 0;

    rows.forEach(row => {
      const matchCollege  = !college      || (row.dataset.college      || '').toLowerCase() === college.toLowerCase();
      const matchBT       = !borrowerType || (row.dataset.borrowerType || '').toLowerCase() === borrowerType.toLowerCase();
      const matchOfficer  = !officer      || (row.dataset.officer      || '').toLowerCase() === officer.toLowerCase();
      const matchMr       = !mr           || (row.dataset.mrLower      || '').toLowerCase() === mr.toLowerCase();
      const matchPtr      = !ptr          || (row.dataset.ptrLower     || '').toLowerCase() === ptr.toLowerCase();
      const matchRelease  = !releaseF     || (row.dataset.release      || '') === releaseF;

      let matchStatus = true;
      if (status) {
        const attrMap = {
          serviceable: 'serviceable', non_serviceable: 'nonServiceable',
          sealed: 'sealed', missing: 'missing', incomplete: 'incomplete',
        };
        const dsKey = attrMap[status];
        matchStatus = dsKey ? row.dataset[dsKey] === '1' : true;
      }

      const matchSearch = rowMatchesSearch(row, search);
      const show = matchSearch && matchCollege && matchBT && matchOfficer && matchMr && matchPtr && matchRelease && matchStatus;
      row.style.display = show ? '' : 'none';
      if (show) visible++;
    });

    const hasFilter = search || college || borrowerType || officer || mr || ptr || releaseF || status;
    const statusLbl = document.getElementById('dm-filter-status-label');
    const clearBtn  = document.getElementById('dm-clear-filters');
    if (statusLbl) statusLbl.style.display = hasFilter ? 'inline-flex' : 'none';
    if (clearBtn)  clearBtn.style.display  = hasFilter ? 'inline-flex' : 'none';
    const vc = document.getElementById('dm-visible-count');
    const tc = document.getElementById('dm-total-count');
    if (vc) vc.textContent = visible;
    if (tc) tc.textContent = rows.length;
  }

  function clearDmFilters() {
    const si = document.getElementById('dm-search');
    if (si) si.value = '';
    ['dm-filter-college', 'dm-filter-borrower-type',
     'dm-filter-officer', 'dm-filter-mr', 'dm-filter-ptr',
     'dm-filter-release', 'dm-filter-status'].forEach(id => {
      const el = document.getElementById(id); if (el) el.selectedIndex = 0;
    });
    populateFilterDropdowns();
    applyDmFilters();
  }

  /* ==================== ADD ROW ==================== */
  function addDmRow() {
    const tbody = document.getElementById('dm-tbody');
    const tr    = document.createElement('tr');
    const newId = 'new_' + Date.now();
    const attrs = {
      'data-row-id': newId, 'data-box': '', 'data-college': '', 'data-college-raw': '',
      'data-person': '', 'data-borrower-type': '', 'data-officer': '', 'data-officer-raw': '',
      'data-device': '', 'data-serial': '', 'data-release': '—',
      'data-mr': '', 'data-mr-lower': '', 'data-ptr': '', 'data-ptr-lower': '',
      'data-serviceable': '0', 'data-non-serviceable': '0',
      'data-sealed': '0', 'data-missing': '0', 'data-incomplete': '0',
    };
    Object.entries(attrs).forEach(([k, v]) => tr.setAttribute(k, v));

    tr.innerHTML = `
      <input type="hidden" name="row_id" value="new"/>
      <td style="text-align:center"><input type="text" name="box_number" class="form-control dm-box-input" placeholder="Box #" style="width:80px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="serial_number" class="form-control dm-serial-input" placeholder="S/N" style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="office_college" class="form-control dm-college-input" placeholder="e.g. CCS" style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="accountable_person" class="form-control dm-person-input" placeholder="Full name" style="width:130px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center">
        <select name="borrower_type" class="form-control dm-borrower-type-select" style="width:90px;text-align:center;margin:0 auto">
          <option value="">— Select —</option>
          <option value="student">Student</option>
          <option value="employee">Employee</option>
        </select>
      </td>
      <td style="text-align:center"><input type="text" name="assigned_mr" class="form-control dm-mr-input" placeholder="M.R. #" style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="accountable_officer" class="form-control dm-officer-input" placeholder="Officer name" style="width:130px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="device" value="Tablet" class="form-control dm-device-input" style="width:90px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="serviceable" value="off"/><input type="checkbox" class="dm-checkbox" data-field="serviceable" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="non_serviceable" value="off"/><input type="checkbox" class="dm-checkbox" data-field="non_serviceable" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="sealed" value="off"/><input type="checkbox" class="dm-checkbox" data-field="sealed" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="missing" value="off"/><input type="checkbox" class="dm-checkbox" data-field="missing" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="incomplete" value="off"/><input type="checkbox" class="dm-checkbox" data-field="incomplete" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="ptr" class="form-control dm-ptr-input" placeholder="PTR #" style="width:100px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><span class="release-status-badge badge-none">—</span></td>
      <td style="text-align:center;color:var(--muted);font-size:12px" class="dm-date-returned">—</td>
      <td style="text-align:center"><textarea name="remarks" class="form-control dm-remarks-input" rows="2" placeholder="Remarks…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto"></textarea></td>
      <td style="text-align:center"><textarea name="issue" class="form-control dm-issue-input" rows="2" placeholder="Issue…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto"></textarea></td>
      <td style="text-align:center;white-space:nowrap">
        <button type="submit" class="btn btn-primary btn-sm">✓ Save</button>
        <button type="button" class="btn btn-danger btn-sm dm-delete-row" style="margin-left:4px">✕</button>
      </td>
    `;
    tbody.appendChild(tr);
    tr.querySelector('input[name="box_number"]').focus();
    applyLockState(tr);
    populateFilterDropdowns();
    applyDmFilters();
  }

  /* ==================== SAVE ALL ROWS (JSON) ==================== */
  const dirtyRows = new Set();

  async function saveAllRows() {
    const form = document.getElementById('dm-form');
    if (!form) return;

    const btn = document.getElementById('saveAllBtn');
    const originalHTML = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = `
        <svg style="width:14px;height:14px;animation:spin .7s linear infinite" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5">
          <path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/>
        </svg>
        Saving…`;
    }

    // Build JSON payload
    const rowsData = [];
    const rows = document.querySelectorAll('#dm-tbody tr[data-row-id]');
    for (const row of rows) {
      const rowId = row.querySelector('input[name="row_id"]')?.value;
      if (!rowId) continue;

      const rowData = {
        row_id: rowId,
        box_number: row.querySelector('input[name="box_number"]')?.value || '',
        serial_number: row.querySelector('input[name="serial_number"]')?.value || '',
        office_college: row.querySelector('input[name="office_college"]')?.value || '',
        accountable_person: row.querySelector('input[name="accountable_person"]')?.value || '',
        borrower_type: row.querySelector('select[name="borrower_type"]')?.value || '',
        assigned_mr: row.querySelector('input[name="assigned_mr"]')?.value || '',
        accountable_officer: row.querySelector('input[name="accountable_officer"]')?.value || '',
        device: row.querySelector('input[name="device"]')?.value || '',
        serviceable: row.querySelector('input[name="serviceable"]')?.nextElementSibling?.checked ? 'on' : 'off',
        non_serviceable: row.querySelector('input[name="non_serviceable"]')?.nextElementSibling?.checked ? 'on' : 'off',
        sealed: row.querySelector('input[name="sealed"]')?.nextElementSibling?.checked ? 'on' : 'off',
        missing: row.querySelector('input[name="missing"]')?.nextElementSibling?.checked ? 'on' : 'off',
        incomplete: row.querySelector('input[name="incomplete"]')?.nextElementSibling?.checked ? 'on' : 'off',
        ptr: row.querySelector('input[name="ptr"]')?.value || '',
        remarks: row.querySelector('textarea[name="remarks"]')?.value || '',
        issue: row.querySelector('textarea[name="issue"]')?.value || '',
      };
      rowsData.push(rowData);
    }

    try {
      const resp = await fetch(form.action, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-CSRFToken': document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '',
          'X-Requested-With': 'XMLHttpRequest',
        },
        body: JSON.stringify({ rows: rowsData, save_all: true }),
        credentials: 'same-origin',
      });

      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

      const result = await resp.json();

      if (result.ok) {
        dirtyRows.clear();
        showToast(`✓ All rows saved (${result.saved} record${result.saved !== 1 ? 's' : ''})`, 'success');
      } else {
        const msg = result.errors?.length
          ? `Saved with errors: ${result.errors.slice(0, 2).join('; ')}`
          : 'Save failed';
        showToast(msg, 'error');
      }
    } catch (err) {
      showToast('Network error: ' + err.message, 'error');
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = originalHTML;
      }
    }
  }

  /* ==================== DELETE ROW ==================== */
  function deleteRow(btn) {
    const row   = btn.closest('tr');
    const rowId = row.dataset.rowId;
    if (rowId && !rowId.startsWith('new_')) {
      if (!confirm('Delete this row?')) return;
      const form = document.createElement('form');
      form.method = 'post';
      form.action = `/device-monitoring/${rowId}/delete/`;
      const csrf = document.createElement('input');
      csrf.type = 'hidden'; csrf.name = 'csrfmiddlewaretoken';
      csrf.value = document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '';
      form.appendChild(csrf);
      document.body.appendChild(form);
      form.submit();
    } else {
      row.remove();
      populateFilterDropdowns();
      applyDmFilters();
    }
  }

  /* ==================== IMPORT MODAL ==================== */
  function openImportModal() {
    const modal = document.getElementById('importModal');
    if (!modal) return;
    modal.style.display = 'flex';
    const fileInput  = document.getElementById('import-file-input');
    const preview    = document.getElementById('import-preview');
    const errorEl    = document.getElementById('import-error');
    const successEl  = document.getElementById('import-success');
    const confirmBtn = document.getElementById('import-confirm-btn');
    if (fileInput)  fileInput.value = '';
    if (preview)    preview.style.display   = 'none';
    if (errorEl)    errorEl.style.display   = 'none';
    if (successEl)  successEl.style.display = 'none';
    if (confirmBtn) confirmBtn.disabled = true;
  }

  function closeImportModal() {
    const modal = document.getElementById('importModal');
    if (modal) modal.style.display = 'none';
  }

  async function confirmImport() {
    const fileInput = document.getElementById('import-file-input');
    const errEl     = document.getElementById('import-error');
    const sucEl     = document.getElementById('import-success');
    const btn       = document.getElementById('import-confirm-btn');
    if (!errEl || !sucEl || !btn || !fileInput) return;
    errEl.style.display = 'none';
    sucEl.style.display = 'none';
    if (!fileInput.files || !fileInput.files[0]) return;

    const formData = new FormData();
    formData.append('excel_file', fileInput.files[0]);
    const csrfToken = document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '';
    formData.append('csrfmiddlewaretoken', csrfToken);

    btn.disabled = true;
    const originalText = btn.innerHTML;
    btn.textContent = 'Importing…';

    try {
      const resp = await fetch('/device-monitoring/import/', {
        method: 'POST', body: formData, credentials: 'same-origin',
      });
      const data = await resp.json();
      if (!data.ok) throw new Error(data.error || 'Import failed');

      let msg = `✓ Import complete: ${data.created} row(s) created, ${data.updated} row(s) updated.`;
      if (data.errors?.length) {
        msg += ` ${data.errors.length} row(s) had errors.`;
        console.warn('Import errors:', data.errors);
      }
      sucEl.textContent = msg;
      sucEl.style.display = 'flex';
      setTimeout(() => { window.location.reload(); }, 1800);
    } catch (err) {
      errEl.textContent = 'Error: ' + err.message;
      errEl.style.display = 'flex';
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
  }

  /* ==================== WEBSOCKET REALTIME ==================== */
  function releaseBadgeHtml(status) {
    if (status === 'Released') return '<span class="release-status-badge badge-released">Released</span>';
    if (status === 'Returned') return '<span class="release-status-badge badge-returned-dm">Returned</span>';
    return '<span class="release-status-badge badge-none">—</span>';
  }

  function handleMessage(data) {
    if (data.type !== 'device_monitoring.update') return;

    const tbody       = document.getElementById('dm-tbody');
    const incomingIds = new Set(data.rows.map(r => String(r.id)));

    tbody.querySelectorAll('tr[data-row-id]').forEach(tr => {
      const id = tr.dataset.rowId;
      if (id && !id.startsWith('new_') && !incomingIds.has(id) && !dirtyRows.has(id)) tr.remove();
    });

    data.rows.forEach(row => {
      const strId    = String(row.id);
      const existing = tbody.querySelector(`tr[data-row-id="${strId}"]`);
      if (!existing || dirtyRows.has(strId)) return;

      const setValue = (name, val) => {
        const el = existing.querySelector(`[name="${name}"]`);
        if (el && document.activeElement !== el) el.value = val || '';
      };
      setValue('box_number',          row.box_number);
      setValue('office_college',      row.office_college);
      setValue('accountable_person',  row.accountable_person);
      setValue('borrower_type',       row.borrower_type);
      setValue('assigned_mr',         row.assigned_mr);
      setValue('accountable_officer', row.accountable_officer);
      setValue('device',              row.device);
      setValue('serial_number',       row.serial_number);
      setValue('ptr',                 row.ptr);
      setValue('remarks',             row.remarks);
      setValue('issue',               row.issue);

      existing.dataset.box          = (row.box_number          || '').toLowerCase();
      existing.dataset.college      = (row.office_college      || '').toLowerCase();
      existing.dataset.collegeRaw   = row.office_college       || '';
      existing.dataset.person       = (row.accountable_person  || '').toLowerCase();
      existing.dataset.borrowerType = (row.borrower_type       || '').toLowerCase();
      existing.dataset.officer      = (row.accountable_officer || '').toLowerCase();
      existing.dataset.officerRaw   = row.accountable_officer  || '';
      existing.dataset.device       = (row.device              || '').toLowerCase();
      existing.dataset.serial       = (row.serial_number       || '').toLowerCase();
      existing.dataset.release      = row.release_status || '—';
      existing.dataset.mr           = row.assigned_mr || '';
      existing.dataset.mrLower      = (row.assigned_mr || '').toLowerCase();
      existing.dataset.ptr          = row.ptr || '';
      existing.dataset.ptrLower     = (row.ptr || '').toLowerCase();
      existing.dataset.serviceable    = row.serviceable     ? '1' : '0';
      existing.dataset.nonServiceable = row.non_serviceable ? '1' : '0';
      existing.dataset.sealed         = row.sealed          ? '1' : '0';
      existing.dataset.missing        = row.missing         ? '1' : '0';
      existing.dataset.incomplete     = row.incomplete      ? '1' : '0';

      const badge = existing.querySelector('.release-status-badge');
      if (badge) badge.outerHTML = releaseBadgeHtml(row.release_status);

      const dateTd = existing.querySelector('.dm-date-returned');
      if (dateTd) dateTd.textContent = (row.date_returned && row.date_returned !== '—') ? row.date_returned : '—';

      const updateCb = (name, val) => {
        const hidden = existing.querySelector(`input[type=hidden][name="${name}"]`);
        const cb = hidden?.nextElementSibling;
        if (cb && !cb.matches(':focus')) {
          if (cb.checked !== val) { cb.checked = val; hidden.value = val ? 'on' : 'off'; }
        }
      };
      updateCb('serviceable',     row.serviceable);
      updateCb('non_serviceable', row.non_serviceable);
      updateCb('sealed',          row.sealed);
      updateCb('missing',         row.missing);
      updateCb('incomplete',      row.incomplete);
      applyLockState(existing);
    });

    sortTableByBoxNumber();
    populateFilterDropdowns();
    applyDmFilters();

    if (data.pending_count !== undefined)
      window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.pending_count }));
    if (data.graduation_warning_count !== undefined)
      window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));
  }

  /* ==================== DRAG-TO-SCROLL ==================== */
  function initDragScroll(container) {
    if (!container) return;
    let isDragging = false, startX = 0, startY = 0,
        scrollLeft = 0, scrollTop = 0, hasDragged = false;
    const DRAG_THRESHOLD = 5;

    container.addEventListener('mousedown', e => {
      if (e.button !== 0) return;
      const tag = e.target.tagName;
      if (['INPUT','TEXTAREA','SELECT','BUTTON','A','LABEL'].includes(tag)) return;
      isDragging = true; hasDragged = false;
      startX = e.pageX - container.offsetLeft;
      startY = e.pageY - container.offsetTop;
      scrollLeft = container.scrollLeft;
      scrollTop  = container.scrollTop;
      container.style.cursor = 'grabbing';
      e.preventDefault();
    });
    document.addEventListener('mousemove', e => {
      if (!isDragging) return;
      const walkX = (e.pageX - container.offsetLeft) - startX;
      const walkY = (e.pageY - container.offsetTop)  - startY;
      if (!hasDragged && (Math.abs(walkX) > DRAG_THRESHOLD || Math.abs(walkY) > DRAG_THRESHOLD))
        hasDragged = true;
      container.scrollLeft = scrollLeft - walkX;
      container.scrollTop  = scrollTop  - walkY;
    });
    document.addEventListener('mouseup', () => {
      if (!isDragging) return;
      isDragging = false;
      container.style.cursor = '';
    });
    container.addEventListener('click', e => {
      if (hasDragged) { e.stopPropagation(); e.preventDefault(); hasDragged = false; }
    }, true);
  }

  /* ==================== EVENT LISTENERS ==================== */
  function attachEventListeners() {
    // Checkbox delegation
    document.addEventListener('change', e => {
      const cb = e.target.closest('.dm-checkbox');
      if (cb?.type === 'checkbox') {
        const field = cb.getAttribute('data-field');
        if (field) { syncCheck(cb); handleDmCheck(cb, field); }
      }
    });

    // Input → dataset sync + dirty tracking
    document.addEventListener('input', e => {
      const row = e.target.closest('tr[data-row-id]');
      if (!row) return;
      if (row.dataset.rowId && !row.dataset.rowId.startsWith('new_'))
        dirtyRows.add(row.dataset.rowId);

      if (e.target.matches('.dm-box-input'))            { row.dataset.box = e.target.value.toLowerCase(); sortTableByBoxNumber(); applyDmFilters(); }
      if (e.target.matches('.dm-college-input'))        { row.dataset.college = e.target.value.toLowerCase(); row.dataset.collegeRaw = e.target.value; populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-person-input'))         { row.dataset.person = e.target.value.toLowerCase(); populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-borrower-type-select')) { row.dataset.borrowerType = e.target.value.toLowerCase(); applyDmFilters(); }
      if (e.target.matches('.dm-officer-input'))        { row.dataset.officer = e.target.value.toLowerCase(); row.dataset.officerRaw = e.target.value; populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-device-input'))         { row.dataset.device = e.target.value.toLowerCase(); applyDmFilters(); }
      if (e.target.matches('.dm-serial-input'))         { row.dataset.serial = e.target.value.toLowerCase(); applyDmFilters(); }
      if (e.target.matches('.dm-mr-input'))             { row.dataset.mr = e.target.value; row.dataset.mrLower = e.target.value.toLowerCase(); populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-ptr-input'))            { row.dataset.ptr = e.target.value; row.dataset.ptrLower = e.target.value.toLowerCase(); populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-remarks-input,.dm-issue-input')) applyDmFilters();
    });

    // Filter controls
    const searchInput = document.getElementById('dm-search');
    if (searchInput) searchInput.addEventListener('input', applyDmFilters);
    ['dm-filter-college','dm-filter-borrower-type','dm-filter-officer',
     'dm-filter-mr','dm-filter-ptr','dm-filter-release','dm-filter-status'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', applyDmFilters);
    });

    const clearBtn = document.getElementById('dm-clear-filters');
    if (clearBtn) clearBtn.addEventListener('click', clearDmFilters);

    const addRowBtn = document.getElementById('addDmRowBtn');
    if (addRowBtn) addRowBtn.addEventListener('click', addDmRow);

    const saveAllBtn = document.getElementById('saveAllBtn');
    if (saveAllBtn) saveAllBtn.addEventListener('click', saveAllRows);

    document.addEventListener('click', e => {
      const deleteBtn = e.target.closest('.dm-delete-row');
      if (deleteBtn) { e.preventDefault(); deleteRow(deleteBtn); }
    });

    const openModalBtn    = document.getElementById('openImportModalBtn');
    const closeModalBtn   = document.getElementById('closeImportModalBtn');
    const cancelImportBtn = document.getElementById('cancelImportBtn');
    const confirmImportBtn = document.getElementById('import-confirm-btn');
    if (openModalBtn)    openModalBtn.addEventListener('click', openImportModal);
    if (closeModalBtn)   closeModalBtn.addEventListener('click', closeImportModal);
    if (cancelImportBtn) cancelImportBtn.addEventListener('click', closeImportModal);
    if (confirmImportBtn) confirmImportBtn.addEventListener('click', confirmImport);

    const fileInput = document.getElementById('import-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', function () {
        const file   = this.files[0];
        const errEl  = document.getElementById('import-error');
        const prevEl = document.getElementById('import-preview');
        const btn    = document.getElementById('import-confirm-btn');
        if (errEl)  errEl.style.display  = 'none';
        if (prevEl) prevEl.style.display = 'none';
        if (btn)    btn.disabled = true;
        if (!file) return;
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
          if (errEl) { errEl.textContent = 'Please select a valid .xlsx or .xls file.'; errEl.style.display = 'flex'; }
          return;
        }
        const previewText = document.getElementById('import-preview-text');
        if (previewText) previewText.textContent = `File: ${file.name}\nSize: ${(file.size/1024).toFixed(1)} KB\nReady to import.`;
        if (prevEl) prevEl.style.display = 'block';
        if (btn)    btn.disabled = false;
      });
    }

    initDragScroll(document.querySelector('.table-container'));
  }

  /* ==================== INITIALISATION ==================== */
  document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('#dm-tbody tr[data-row-id]').forEach(row => applyLockState(row));
    sortTableByBoxNumber();
    populateFilterDropdowns();
    applyDmFilters();
    attachEventListeners();

    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined' && InvSysRT.connect)
      InvSysRT.connect('/ws/device-monitoring/', handleMessage, indicator);
  });
})();