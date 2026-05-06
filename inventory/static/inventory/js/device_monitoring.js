/**
 * device_monitoring.js — Release/Return badge with inline dropdown edit
 *
 * FIX: release_status was not being saved on new rows.
 *
 * Root causes and changes:
 *
 * 1. addDmRow() was firing an immediate fetch with a hardcoded
 *    `release_status: ''` payload. This locked in an empty value in DB
 *    before the user could pick anything. FIXED: removed the immediate
 *    fetch from addDmRow(). The row is now created in DB only when the
 *    user blurs a field — same behaviour as all other columns.
 *
 * 2. scheduleAutoSave() skipped rows whose id starts with 'new_'. FIXED:
 *    it now calls _saveNewRow(tr) for unsaved rows, which does the
 *    create + id-swap using extractRowPayload() so ALL current field
 *    values (including release_status) are captured at save time.
 *
 * 3. saveRow() had an early return for 'new_' rows, so the badge dropdown
 *    change handler (saveAndRestore) was silently dropping the save if the
 *    row hadn't been committed yet. FIXED: saveRow now delegates to
 *    _saveNewRow() for 'new_' rows instead of returning early.
 *
 * 4. saveAndRestore() in attachReleaseEditListener was only calling
 *    saveRow(tr). That is now correct because saveRow handles new_ rows too.
 *
 * No changes to views.py are needed.
 */

(function () {
  'use strict';

  /* ==================== CONSTANTS ==================== */
  const ROW_HEIGHT     = 64;
  const OVERSCAN       = 10;
  const VISIBLE_BUFFER = 15;
  const AUTOSAVE_DELAY = 800;

  // Per-row debounce timers (existing rows only — new rows use _saveNewRow)
  const _saveTimers = new Map();

  // Track new rows that are currently being saved (prevent double-fire)
  const _savingNew  = new Set();

  async function pollExportTask(taskId) {
    const hide = () => {
      const overlay = document.getElementById('invsys-loading-overlay');
      if (overlay) overlay.classList.remove('is-active');
    };
    try {
      const resp = await fetch(`/export-task-status/${taskId}/`);
      const data = await resp.json();
      if (data.state === 'SUCCESS') {
        window.location.href = `/download-export/${data.token}/`;
        hide();
      } else if (data.state === 'FAILURE') {
        alert('Export failed.');
        hide();
      } else {
        setTimeout(() => pollExportTask(taskId), 2000);
      }
    } catch (err) {
      hide();
    }
  }

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
    }, 3000);
  }

  /* ==================== DEBOUNCE ==================== */
  function debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  }

  function normalizeDmRow(r) {
    const o = { ...r };
    o.date_returned_display = o.date_returned_display ?? o.date_returned ?? '—';
    return o;
  }

  /* ==================== STATE ==================== */
  let allRows      = [];
  let filteredRows = [];
  const dirtyRows  = new Set();

  let scrollTop       = 0;
  let containerHeight = 600;
  let startIdx = 0, endIdx = 0;

  let tbody, scrollContainer, topSpacer, bottomSpacer;

  /* ==================== CSRF ==================== */
  function getCsrf() {
    return document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '';
  }

  /* ==================== SAVING INDICATOR ==================== */
  function _setRowStatus(tr, status) {
    let badge = tr.querySelector('.dm-row-status');
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'dm-row-status';
      badge.style.cssText = `
        display:inline-block;font-size:10px;font-weight:600;
        padding:2px 6px;border-radius:4px;margin-left:4px;
        transition:opacity .3s;`;
      const lastTd = tr.querySelector('td:last-child');
      if (lastTd) lastTd.appendChild(badge);
    }
    if (status === 'saving') {
      badge.textContent = 'Saving…';
      badge.style.background = 'rgba(255,255,255,.12)';
      badge.style.color = 'var(--muted, #888)';
      badge.style.opacity = '1';
    } else if (status === 'saved') {
      badge.textContent = '✓ Saved';
      badge.style.background = 'rgba(0,229,160,.15)';
      badge.style.color = '#00e5a0';
      badge.style.opacity = '1';
      setTimeout(() => { badge.style.opacity = '0'; }, 2000);
    } else if (status === 'error') {
      badge.textContent = '✕ Error';
      badge.style.background = 'rgba(255,76,76,.15)';
      badge.style.color = '#ff4c4c';
      badge.style.opacity = '1';
    }
  }

  /* ==================== CHECKBOX HELPERS ==================== */
  window.syncCheck = function (cb) {
    if (cb.previousElementSibling) cb.previousElementSibling.value = cb.checked ? 'on' : 'off';
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
    if (!row) return;
    const c = getChecksInRow(row);
    if (cb.checked) {
      if (['non_serviceable', 'missing', 'incomplete'].includes(field)) {
        Object.keys(c).forEach(k => {
          if (k !== field) { if (c[k].cb) c[k].cb.checked = false; if (c[k].hidden) c[k].hidden.value = 'off'; }
        });
      }
      if (field === 'serviceable' || field === 'sealed') {
        ['non_serviceable', 'missing', 'incomplete'].forEach(k => {
          if (c[k].cb) c[k].cb.checked = false; if (c[k].hidden) c[k].hidden.value = 'off';
        });
      }
    }
    applyLockState(row);
    markDirtyFromRow(row);
    scheduleAutoSave(row);
  };

  function applyLockState(row) {
    const c = getChecksInRow(row);
    if (!c.serviceable.cb) return;
    const exclusiveOn = c.non_serviceable.cb?.checked || c.missing.cb?.checked || c.incomplete.cb?.checked;
    const safeOn      = c.serviceable.cb?.checked || c.sealed.cb?.checked;
    if (exclusiveOn) {
      Object.values(c).forEach(x => { if (x.cb) x.cb.disabled = !x.cb.checked; });
    } else if (safeOn) {
      ['non_serviceable', 'missing', 'incomplete'].forEach(k => { if (c[k].cb) c[k].cb.disabled = true; });
      if (c.serviceable.cb) c.serviceable.cb.disabled = false;
      if (c.sealed.cb)      c.sealed.cb.disabled      = false;
    } else {
      Object.values(c).forEach(x => { if (x.cb) x.cb.disabled = false; });
    }
  }

  function markDirtyFromRow(row) {
    const rowId = row?.dataset?.rowId;
    if (rowId && !rowId.startsWith('new_')) dirtyRows.add(rowId);
  }

  /* ==================== AUTO-SAVE SCHEDULER ==================== */
  /**
   * FIX: now handles new_ rows by calling _saveNewRow() instead of skipping.
   */
  function scheduleAutoSave(tr) {
    const rowId = tr?.dataset?.rowId;
    if (!rowId) return;

    // ── New (unsaved) row — create in DB ─────────────────────────────────────
    if (rowId.startsWith('new_')) {
      if (_savingNew.has(rowId)) return;   // already in flight
      if (_saveTimers.has(rowId)) clearTimeout(_saveTimers.get(rowId));
      _setRowStatus(tr, 'saving');
      const timer = setTimeout(() => {
        _saveTimers.delete(rowId);
        _saveNewRow(tr);
      }, AUTOSAVE_DELAY);
      _saveTimers.set(rowId, timer);
      return;
    }

    // ── Existing row — debounced update ──────────────────────────────────────
    if (_saveTimers.has(rowId)) clearTimeout(_saveTimers.get(rowId));
    _setRowStatus(tr, 'saving');
    const timer = setTimeout(() => {
      _saveTimers.delete(rowId);
      saveRow(tr);
    }, AUTOSAVE_DELAY);
    _saveTimers.set(rowId, timer);
  }

  /* ==================== SAVE NEW ROW (CREATE) ==================== */
  /**
   * FIX: reads ALL current field values via extractRowPayload() — including
   * the release_status hidden input — so whatever the user has set is captured.
   * Previously addDmRow() sent a hardcoded payload with release_status: ''.
   */
  async function _saveNewRow(tr) {
    const clientId = tr?.dataset?.rowId;
    if (!clientId || !clientId.startsWith('new_')) return;
    if (_savingNew.has(clientId)) return;

    _savingNew.add(clientId);
    harvestRowEdits(tr, clientId);
    const payload = extractRowPayload(tr);   // ← reads release_status from hidden input

    const form = document.getElementById('dm-form');
    if (!form) { _savingNew.delete(clientId); return; }

    _setRowStatus(tr, 'saving');

    try {
      const resp = await fetch(form.action, {
        method:  'POST',
        headers: {
          'Content-Type':     'application/json',
          'X-CSRFToken':      getCsrf(),
          'X-Requested-With': 'XMLHttpRequest',
        },
        body:        JSON.stringify({ rows: [payload], save_all: false }),
        credentials: 'same-origin',
      });

      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const result = await resp.json();

      if (result.ok && result.new_ids?.length > 0) {
        _swapRowId(tr, clientId, String(result.new_ids[0]));
        _setRowStatus(tr, 'saved');
      } else {
        _setRowStatus(tr, 'error');
        showToast('Could not save new row', 'error');
      }
    } catch (err) {
      _setRowStatus(tr, 'error');
      showToast('Row save failed: ' + err.message, 'error');
    } finally {
      _savingNew.delete(clientId);
    }
  }

  /* ==================== RELEASE BADGE + INLINE EDIT ==================== */
  function getReleaseBadgeHtml(value) {
    const displayValue = value || '—';
    let badgeClass = 'badge-none';
    if (displayValue === 'Released') badgeClass = 'badge-released';
    if (displayValue === 'Returned') badgeClass = 'badge-returned-dm';
    return `<span class="release-status-badge ${badgeClass}" data-release-value="${displayValue}" style="cursor:pointer;">${_esc(displayValue)}</span>`;
  }

  function createReleaseDropdown(currentValue) {
    const select = document.createElement('select');
    select.className = 'form-control dm-release-select temp-release-select';
    select.style.cssText = 'width:100px;text-align:center;margin:0 auto;padding:2px;';
    const emptyOpt    = document.createElement('option'); emptyOpt.value    = ''; emptyOpt.textContent    = '—';
    const releasedOpt = document.createElement('option'); releasedOpt.value = 'Released'; releasedOpt.textContent = 'Released';
    const returnedOpt = document.createElement('option'); returnedOpt.value = 'Returned'; returnedOpt.textContent = 'Returned';
    select.appendChild(emptyOpt);
    select.appendChild(releasedOpt);
    select.appendChild(returnedOpt);
    if (currentValue === 'Released')      select.value = 'Released';
    else if (currentValue === 'Returned') select.value = 'Returned';
    else                                  select.value = '';
    return select;
  }

  function attachReleaseEditListener(container) {
    container.addEventListener('click', (e) => {
      const badge = e.target.closest('.release-status-badge');
      if (!badge) return;
      e.stopPropagation();

      const td    = badge.closest('td');
      const tr    = td.closest('tr[data-row-id]');
      const rowId = tr?.dataset.rowId;
      if (!rowId) return;

      const currentValue = badge.getAttribute('data-release-value') === '—'
        ? '' : badge.getAttribute('data-release-value');
      const dropdown = createReleaseDropdown(currentValue);
      td.innerHTML = '';
      td.appendChild(dropdown);
      dropdown.focus();

      const saveAndRestore = async () => {
        const newValue = dropdown.value;

        // Always update in-memory + hidden input regardless of whether changed
        const dataRow = allRows.find(r => String(r.id) === rowId);
        if (dataRow) dataRow.release_status = newValue;
        tr.dataset.release = newValue || '—';

        const hiddenInput = tr.querySelector('input[type=hidden][name="release_status"]');
        if (hiddenInput) hiddenInput.value = newValue;

        // Restore badge immediately so the user sees the new value
        td.innerHTML = getReleaseBadgeHtml(newValue);

        // Only save if the value actually changed
        if (newValue !== currentValue) {
          markDirtyFromRow(tr);
          // FIX: saveRow now handles new_ rows too (delegates to _saveNewRow)
          await saveRow(tr);
        }
      };

      dropdown.addEventListener('change', saveAndRestore);
      dropdown.addEventListener('blur', () => {
        setTimeout(() => {
          if (document.activeElement !== dropdown) saveAndRestore();
        }, 100);
      });
    });
  }

  /* ==================== BUILD ROW HTML ==================== */
  function _esc(str) {
    return String(str || '')
      .replace(/&/g, '&amp;').replace(/"/g, '&quot;')
      .replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function buildRowHtml(row) {
    const releaseBadgeHtml = getReleaseBadgeHtml(row.release_status);
    return `<input type="hidden" name="row_id" value="${row.id}"/>
      <td style="text-align:center"><input type="text" name="box_number" value="${_esc((row.box_number || '').replace(/\.0$/, ''))}" class="form-control dm-box-input" placeholder="Box #" style="width:80px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="serial_number" value="${_esc(row.serial_number)}" class="form-control dm-serial-input" placeholder="S/N" style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="office_college" value="${_esc(row.office_college)}" class="form-control dm-college-input" placeholder="e.g. CCS" style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="accountable_person" value="${_esc(row.accountable_person)}" class="form-control dm-person-input" placeholder="Full name" style="width:130px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center">
        <select name="borrower_type" class="form-control dm-borrower-type-select" style="width:90px;text-align:center;margin:0 auto">
          <option value="">— Select —</option>
          <option value="student"  ${row.borrower_type === 'student'  ? 'selected' : ''}>Student</option>
          <option value="employee" ${row.borrower_type === 'employee' ? 'selected' : ''}>Employee</option>
        </select>
      </td>
      <td style="text-align:center"><input type="text" name="assigned_mr" value="${_esc(row.assigned_mr)}" class="form-control dm-mr-input" placeholder="M.R." style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="accountable_officer" value="${_esc(row.accountable_officer)}" class="form-control dm-officer-input" placeholder="Officer name" style="width:130px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="device" value="${_esc(row.device)}" class="form-control dm-device-input" style="width:90px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="serviceable" value="${row.serviceable ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="serviceable" ${row.serviceable ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="non_serviceable" value="${row.non_serviceable ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="non_serviceable" ${row.non_serviceable ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="sealed" value="${row.sealed ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="sealed" ${row.sealed ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="missing" value="${row.missing ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="missing" ${row.missing ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="incomplete" value="${row.incomplete ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="incomplete" ${row.incomplete ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="ptr" value="${_esc(row.ptr)}" class="form-control dm-ptr-input" placeholder="PTR" style="width:100px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center">
        <input type="hidden" name="release_status" value="${_esc(row.release_status || '')}"/>
        ${releaseBadgeHtml}
      </td>
      <td style="text-align:center;color:var(--muted);font-size:12px" class="dm-date-returned">${_esc(row.date_returned_display || '—')}</td>
      <td style="text-align:center"><textarea name="remarks" class="form-control dm-remarks-input" rows="2" placeholder="Remarks…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto">${_esc(row.remarks)}</textarea></td>
      <td style="text-align:center"><textarea name="issue" class="form-control dm-issue-input" rows="2" placeholder="Issue…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto">${_esc(row.issue)}</textarea></td>
      <td style="text-align:center;white-space:nowrap">
        <button type="button" class="btn btn-danger btn-sm dm-delete-row">✕</button>
      </td>`;
  }

  /* ==================== VIRTUAL SCROLL ENGINE ==================== */
  const rowIdToElement = new Map();
  let focusedRowId = null;

  function getVisibleRange() {
    const visibleStart = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN);
    const visibleCount = Math.ceil(containerHeight / ROW_HEIGHT) + OVERSCAN * 2 + VISIBLE_BUFFER;
    const visibleEnd   = Math.min(filteredRows.length, visibleStart + visibleCount);
    return [visibleStart, visibleEnd];
  }

  function renderVisibleRows() {
    if (!tbody) return;
    const [newStart, newEnd] = getVisibleRange();

    topSpacer.style.height    = (newStart * ROW_HEIGHT) + 'px';
    bottomSpacer.style.height = (Math.max(0, filteredRows.length - newEnd) * ROW_HEIGHT) + 'px';

    const neededIds = new Set();
    for (let i = newStart; i < newEnd; i++) neededIds.add(String(filteredRows[i].id));
    if (focusedRowId) neededIds.add(focusedRowId);
    dirtyRows.forEach(id => neededIds.add(id));
    _saveTimers.forEach((_, id) => neededIds.add(id));
    // Always keep unsaved new rows in DOM
    allRows.forEach(r => { if (String(r.id).startsWith('new_')) neededIds.add(String(r.id)); });

    rowIdToElement.forEach((tr, id) => {
      if (!neededIds.has(id)) {
        harvestRowEdits(tr, id);
        tr.remove();
        rowIdToElement.delete(id);
      }
    });

    const fragment = document.createDocumentFragment();
    for (let i = newStart; i < newEnd; i++) {
      const row   = filteredRows[i];
      const strId = String(row.id);
      if (rowIdToElement.has(strId)) continue;

      const tr = document.createElement('tr');
      tr.dataset.rowId        = strId;
      tr.dataset.box          = (row.box_number          || '').toLowerCase();
      tr.dataset.college      = (row.office_college      || '').toLowerCase();
      tr.dataset.collegeRaw   =  row.office_college      || '';
      tr.dataset.person       = (row.accountable_person  || '').toLowerCase();
      tr.dataset.borrowerType = (row.borrower_type       || '').toLowerCase();
      tr.dataset.officer      = (row.accountable_officer || '').toLowerCase();
      tr.dataset.officerRaw   =  row.accountable_officer || '';
      tr.dataset.device       = (row.device              || '').toLowerCase();
      tr.dataset.serial       = (row.serial_number       || '').toLowerCase();
      tr.dataset.release      =  row.release_status      || '—';
      tr.dataset.mr           =  row.assigned_mr         || '';
      tr.dataset.mrLower      = (row.assigned_mr         || '').toLowerCase();
      tr.dataset.ptr          =  row.ptr                 || '';
      tr.dataset.ptrLower     = (row.ptr                 || '').toLowerCase();
      tr.dataset.serviceable    = row.serviceable     ? '1' : '0';
      tr.dataset.nonServiceable = row.non_serviceable ? '1' : '0';
      tr.dataset.sealed         = row.sealed          ? '1' : '0';
      tr.dataset.missing        = row.missing         ? '1' : '0';
      tr.dataset.incomplete     = row.incomplete      ? '1' : '0';
      tr.innerHTML = buildRowHtml(row);
      applyLockState(tr);
      rowIdToElement.set(strId, tr);
      fragment.appendChild(tr);
    }

    tbody.insertBefore(fragment, bottomSpacer);
    startIdx = newStart;
    endIdx   = newEnd;

    const vc = document.getElementById('dm-visible-count');
    const tc = document.getElementById('dm-total-count');
    if (vc) vc.textContent = filteredRows.length;
    if (tc) tc.textContent = allRows.length;
  }

  function harvestRowEdits(tr, id) {
    const dataRow = allRows.find(r => String(r.id) === String(id));
    if (!dataRow) return;
    const g = (name) => tr.querySelector(`[name="${name}"]`)?.value ?? dataRow[name];
    dataRow.box_number          = g('box_number');
    dataRow.serial_number       = g('serial_number');
    dataRow.office_college      = g('office_college');
    dataRow.accountable_person  = g('accountable_person');
    dataRow.borrower_type       = tr.querySelector('select[name="borrower_type"]')?.value ?? dataRow.borrower_type;
    dataRow.assigned_mr         = g('assigned_mr');
    dataRow.accountable_officer = g('accountable_officer');
    dataRow.device              = g('device');
    dataRow.ptr                 = g('ptr');
    dataRow.remarks             = tr.querySelector('textarea[name="remarks"]')?.value ?? dataRow.remarks;
    dataRow.issue               = tr.querySelector('textarea[name="issue"]')?.value ?? dataRow.issue;
    // Always harvest release_status from the hidden input
    const releaseHidden = tr.querySelector('input[type=hidden][name="release_status"]');
    if (releaseHidden) dataRow.release_status = releaseHidden.value;
    const cbState = (name) => tr.querySelector(`input[type=hidden][name="${name}"]`)?.value === 'on';
    dataRow.serviceable     = cbState('serviceable');
    dataRow.non_serviceable = cbState('non_serviceable');
    dataRow.sealed          = cbState('sealed');
    dataRow.missing         = cbState('missing');
    dataRow.incomplete      = cbState('incomplete');
  }

  /* ==================== EXTRACT ROW PAYLOAD ==================== */
  function extractRowPayload(tr) {
    const rawId = tr.dataset.rowId || '';
    const isNew = rawId.startsWith('new_');
    const g = (name) => tr.querySelector(`[name="${name}"]`)?.value ?? '';
    const cbState = (name) => tr.querySelector(`input[type=hidden][name="${name}"]`)?.value === 'on';
    // Read release_status directly from the hidden input — catches any value
    // the user set via the badge dropdown before the save fires
    const releaseStatus = tr.querySelector('input[type=hidden][name="release_status"]')?.value || '';

    return {
      row_id:              isNew ? 'new' : rawId,
      _client_id:          rawId,
      box_number:          g('box_number'),
      serial_number:       g('serial_number'),
      office_college:      g('office_college'),
      accountable_person:  g('accountable_person'),
      borrower_type:       tr.querySelector('select[name="borrower_type"]')?.value ?? '',
      assigned_mr:         g('assigned_mr'),
      accountable_officer: g('accountable_officer'),
      device:              g('device') || 'Tablet',
      serviceable:         cbState('serviceable')     ? 'on' : 'off',
      non_serviceable:     cbState('non_serviceable') ? 'on' : 'off',
      sealed:              cbState('sealed')          ? 'on' : 'off',
      missing:             cbState('missing')         ? 'on' : 'off',
      incomplete:          cbState('incomplete')      ? 'on' : 'off',
      ptr:                 g('ptr'),
      remarks:             tr.querySelector('textarea[name="remarks"]')?.value ?? '',
      issue:               tr.querySelector('textarea[name="issue"]')?.value ?? '',
      release_status:      releaseStatus,
    };
  }

  /* ==================== FILTER / SEARCH ==================== */
  function applyFilters() {
    const search       = (document.getElementById('dm-search')?.value || '').toLowerCase().trim();
    const college      = document.getElementById('dm-filter-college')?.value     || '';
    const borrowerType = document.getElementById('dm-filter-borrower-type')?.value || '';
    const officer      = document.getElementById('dm-filter-officer')?.value      || '';
    const mr           = document.getElementById('dm-filter-mr')?.value           || '';
    const ptr          = document.getElementById('dm-filter-ptr')?.value          || '';
    const releaseF     = document.getElementById('dm-filter-release')?.value      || '';
    const status       = document.getElementById('dm-filter-status')?.value       || '';

    filteredRows = allRows.filter(row => {
      if (search) {
        const hay = [row.box_number, row.serial_number, row.office_college,
                     row.accountable_person, row.accountable_officer].join(' ').toLowerCase();
        if (!hay.includes(search)) return false;
      }
      if (college      && (row.office_college      || '').toLowerCase() !== college.toLowerCase())      return false;
      if (borrowerType && (row.borrower_type       || '').toLowerCase() !== borrowerType.toLowerCase()) return false;
      if (officer      && (row.accountable_officer || '').toLowerCase() !== officer.toLowerCase())      return false;
      if (mr           && (row.assigned_mr         || '').toLowerCase() !== mr.toLowerCase())           return false;
      if (ptr          && (row.ptr                 || '').toLowerCase() !== ptr.toLowerCase())           return false;
      if (releaseF     && (row.release_status      || '') !== releaseF)                                 return false;
      if (status) {
        const map = { serviceable: 'serviceable', non_serviceable: 'non_serviceable',
                      sealed: 'sealed', missing: 'missing', incomplete: 'incomplete' };
        if (map[status] && !row[map[status]]) return false;
      }
      return true;
    });

    rowIdToElement.forEach(tr => { harvestRowEdits(tr, tr.dataset.rowId); tr.remove(); });
    rowIdToElement.clear();

    if (scrollContainer) scrollContainer.scrollTop = 0;
    scrollTop = 0;
    renderVisibleRows();

    const hasFilter = search || college || borrowerType || officer || mr || ptr || releaseF || status;
    const statusLbl = document.getElementById('dm-filter-status-label');
    const clearBtn  = document.getElementById('dm-clear-filters');
    if (statusLbl) statusLbl.style.display = hasFilter ? 'inline-flex' : 'none';
    if (clearBtn)  clearBtn.style.display  = hasFilter ? 'inline-flex' : 'none';
  }

  const debouncedFilter = debounce(applyFilters, 200);

  /* ==================== FILTER DROPDOWNS ==================== */
  function populateFilterDropdowns() {
    const collegeSelect = document.getElementById('dm-filter-college');
    const officerSelect = document.getElementById('dm-filter-officer');
    const mrSelect      = document.getElementById('dm-filter-mr');
    const ptrSelect     = document.getElementById('dm-filter-ptr');
    if (!collegeSelect) return;

    const saved = {
      college: collegeSelect.value,
      officer: officerSelect?.value || '',
      mr:      mrSelect?.value      || '',
      ptr:     ptrSelect?.value     || '',
    };

    const collegeSet = new Set(), officerSet = new Set(), mrSet = new Set(), ptrSet = new Set();
    allRows.forEach(r => {
      if (r.office_college)      collegeSet.add(r.office_college);
      if (r.accountable_officer) officerSet.add(r.accountable_officer);
      if (r.assigned_mr)         mrSet.add(r.assigned_mr);
      if (r.ptr)                 ptrSet.add(r.ptr);
    });

    const fill = (sel, set, savedVal) => {
      if (!sel) return;
      while (sel.options.length > 1) sel.remove(1);
      [...set].sort((a, b) => a.localeCompare(b)).forEach(v => {
        const o = document.createElement('option');
        o.value = v; o.textContent = v;
        if (v === savedVal) o.selected = true;
        sel.appendChild(o);
      });
    };
    fill(collegeSelect, collegeSet, saved.college);
    fill(officerSelect, officerSet, saved.officer);
    fill(mrSelect,      mrSet,      saved.mr);
    fill(ptrSelect,     ptrSet,     saved.ptr);
  }

  /* ==================== BOX NUMBER SORT ==================== */
  function sortByBoxNumber(rows) {
    return rows.slice().sort((a, b) => {
      const numA = parseInt((a.box_number || '').match(/(\d+)/)?.[1] ?? 'Infinity', 10);
      const numB = parseInt((b.box_number || '').match(/(\d+)/)?.[1] ?? 'Infinity', 10);
      return numA - numB;
    });
  }

  /* ==================== INITIAL RENDER ==================== */
  function initVirtualScroll(rowsData) {
    allRows      = sortByBoxNumber(rowsData);
    filteredRows = allRows;

    topSpacer    = document.createElement('tr');
    bottomSpacer = document.createElement('tr');
    topSpacer.style.height    = '0px';
    bottomSpacer.style.height = (allRows.length * ROW_HEIGHT) + 'px';
    tbody.appendChild(topSpacer);
    tbody.appendChild(bottomSpacer);

    scrollContainer = document.querySelector('.table-container');
    if (!scrollContainer) return;

    containerHeight = scrollContainer.clientHeight || 600;

    scrollContainer.addEventListener('scroll', () => {
      scrollTop = scrollContainer.scrollTop;
      renderVisibleRows();
    }, { passive: true });

    if (window.ResizeObserver) {
      new ResizeObserver(entries => {
        containerHeight = entries[0].contentRect.height;
        renderVisibleRows();
      }).observe(scrollContainer);
    }

    renderVisibleRows();
    populateFilterDropdowns();

    const tc = document.getElementById('dm-total-count');
    if (tc) tc.textContent = allRows.length;
  }

  /* ==================== ADD ROW ==================== */
  /**
   * FIX: removed the immediate fetch that was hardcoding release_status: ''.
   * The row is now only saved to DB when the user actually fills a field and
   * blurs away from it — at that point, scheduleAutoSave → _saveNewRow reads
   * all current values including the release_status hidden input.
   */
  function addDmRow() {
    const newClientId = 'new_' + Date.now();
    const newRow = {
      id: newClientId, box_number: '', serial_number: '', office_college: '',
      accountable_person: '', borrower_type: '', assigned_mr: '',
      accountable_officer: '', device: 'Tablet', serviceable: false,
      non_serviceable: false, sealed: false, missing: false, incomplete: false,
      ptr: '', remarks: '', issue: '', release_status: '',
      date_returned_display: '—',
    };

    allRows.push(newRow);
    filteredRows.push(newRow);

    const tr = document.createElement('tr');
    tr.dataset.rowId = newClientId;
    tr.innerHTML = buildRowHtml(newRow);
    applyLockState(tr);
    tbody.insertBefore(tr, bottomSpacer);
    rowIdToElement.set(newClientId, tr);

    if (scrollContainer) scrollContainer.scrollTop = scrollContainer.scrollHeight;
    tr.querySelector('input[name="box_number"]')?.focus();

    const tc = document.getElementById('dm-total-count');
    if (tc) tc.textContent = allRows.length;

    // No immediate save here — _saveNewRow fires via scheduleAutoSave on blur
  }

  /* ==================== SAVE EXISTING ROW ==================== */
  /**
   * FIX: removed `if (clientId.startsWith('new_')) return;` — instead
   * delegates to _saveNewRow so the release badge dropdown's saveAndRestore()
   * correctly saves even before the first blur has committed the row.
   */
  async function saveRow(tr) {
    if (!tr) return;
    const clientId = tr.dataset.rowId;

    // FIX: delegate to _saveNewRow instead of silently returning
    if (clientId.startsWith('new_')) {
      return _saveNewRow(tr);
    }

    harvestRowEdits(tr, clientId);
    const payload = extractRowPayload(tr);

    const form = document.getElementById('dm-form');
    if (!form) return;

    _setRowStatus(tr, 'saving');

    try {
      const resp = await fetch(form.action, {
        method:  'POST',
        headers: {
          'Content-Type':     'application/json',
          'X-CSRFToken':      getCsrf(),
          'X-Requested-With': 'XMLHttpRequest',
        },
        body:        JSON.stringify({ rows: [payload], save_all: false }),
        credentials: 'same-origin',
      });

      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const result = await resp.json();

      if (!result.ok) throw new Error(result.errors?.[0] || 'Save failed');

      dirtyRows.delete(clientId);
      _setRowStatus(tr, 'saved');
    } catch (err) {
      _setRowStatus(tr, 'error');
      showToast('Auto-save error: ' + err.message, 'error');
    }
  }

  /* ==================== ID SWAP HELPER ==================== */
  function _swapRowId(tr, oldId, newId) {
    tr.dataset.rowId = newId;
    const hidden = tr.querySelector('input[type=hidden][name="row_id"]');
    if (hidden) hidden.value = newId;

    rowIdToElement.delete(oldId);
    rowIdToElement.set(newId, tr);

    const dataRow = allRows.find(r => String(r.id) === oldId);
    if (dataRow) dataRow.id = newId;

    const fRow = filteredRows.find(r => String(r.id) === oldId);
    if (fRow) fRow.id = newId;

    dirtyRows.delete(oldId);
    _saveTimers.delete(oldId);
    _savingNew.delete(oldId);
  }

  /* ==================== DELETE ROW ==================== */
  function deleteRow(btn) {
    const tr    = btn.closest('tr');
    const rowId = tr?.dataset?.rowId;
    if (!rowId) return;

    if (rowId.startsWith('new_')) {
      allRows      = allRows.filter(r => String(r.id) !== rowId);
      filteredRows = filteredRows.filter(r => String(r.id) !== rowId);
      rowIdToElement.delete(rowId);
      if (_saveTimers.has(rowId)) { clearTimeout(_saveTimers.get(rowId)); _saveTimers.delete(rowId); }
      _savingNew.delete(rowId);
      tr.remove();
      const tc = document.getElementById('dm-total-count');
      if (tc) tc.textContent = allRows.length;
      return;
    }

    if (!confirm('Delete this row? This cannot be undone.')) return;

    const overlay = document.getElementById('invsys-loading-overlay');
    const label   = overlay?.querySelector('.loader-label');
    if (overlay) {
      if (label) label.textContent = 'Deleting…';
      overlay.classList.add('is-active');
      overlay.setAttribute('aria-hidden', 'false');
    }

    if (_saveTimers.has(rowId)) { clearTimeout(_saveTimers.get(rowId)); _saveTimers.delete(rowId); }

    fetch(`/device-monitoring/${rowId}/delete/`, {
      method:      'POST',
      credentials: 'same-origin',
      headers:     { 'X-CSRFToken': getCsrf(), 'X-Requested-With': 'XMLHttpRequest' },
    })
      .then(() => {
        allRows      = allRows.filter(r => String(r.id) !== rowId);
        filteredRows = filteredRows.filter(r => String(r.id) !== rowId);
        rowIdToElement.delete(rowId);
        dirtyRows.delete(rowId);
        tr.remove();
        if (bottomSpacer) {
          bottomSpacer.style.height = (Math.max(0, filteredRows.length - endIdx) * ROW_HEIGHT) + 'px';
        }
        const tc = document.getElementById('dm-total-count');
        if (tc) tc.textContent = allRows.length;
        showToast('Row deleted', 'success');
      })
      .catch(err => showToast('Delete failed: ' + err.message, 'error'))
      .finally(() => {
        if (overlay) {
          if (label) label.textContent = 'Loading…';
          overlay.classList.remove('is-active');
          overlay.setAttribute('aria-hidden', 'true');
        }
      });
  }

  /* ==================== WEBSOCKET REALTIME ==================== */
  function handleMessage(data) {
    if (data.type !== 'device_monitoring.update') return;

    const incoming = new Map(data.rows.map(r => [String(r.id), r]));

    for (let i = 0; i < allRows.length; i++) {
      const id = String(allRows[i].id);
      if (incoming.has(id) && !dirtyRows.has(id) && !_saveTimers.has(id)) {
        const inc = normalizeDmRow(incoming.get(id));
        allRows[i] = { ...allRows[i], ...inc };
        const tr = rowIdToElement.get(id);
        if (tr && id !== focusedRowId) {
          const badge  = tr.querySelector('.release-status-badge');
          const hidden = tr.querySelector('input[type=hidden][name="release_status"]');
          if (inc.release_status !== undefined) {
            const newValue = inc.release_status || '';
            if (hidden) hidden.value = newValue;
            if (badge) {
              const cls = newValue === 'Released' ? 'badge-released'
                        : newValue === 'Returned'  ? 'badge-returned-dm' : 'badge-none';
              badge.className = `release-status-badge ${cls}`;
              badge.textContent = newValue || '—';
              badge.setAttribute('data-release-value', newValue || '—');
              tr.dataset.release = newValue || '—';
            }
          }
          const dateTd = tr.querySelector('.dm-date-returned');
          if (dateTd) dateTd.textContent = inc.date_returned_display || '—';
        }
      }
    }

    if (typeof data.pending_count === 'number')
      window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.pending_count }));
    if (typeof data.graduation_warning_count === 'number')
      window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));
  }

  /* ==================== IMPORT MODAL ==================== */
  function openImportModal() {
    const modal = document.getElementById('importModal');
    if (!modal) return;
    modal.style.display = 'flex';
    _resetImportModal();
  }
  function _resetImportModal() {
    ['import-file-input','import-preview','import-error','import-success','import-progress-wrap'].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (id === 'import-file-input') el.value = '';
      else el.style.display = 'none';
    });
    const btn = document.getElementById('import-confirm-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Import'; }
  }
  function closeImportModal() {
    const modal = document.getElementById('importModal');
    if (modal) modal.style.display = 'none';
    _stopPolling();
  }

  let _pollTimer = null, _currentTask = null, _progressTick = 10;
  function _stopPolling() { clearInterval(_pollTimer); _pollTimer = null; _currentTask = null; }
  function _setImportProgress(pct, label) {
    const wrap = document.getElementById('import-progress-wrap');
    const bar  = document.getElementById('import-progress-bar');
    const lbl  = document.getElementById('import-progress-label');
    if (wrap) wrap.style.display = 'block';
    if (bar)  bar.style.width = pct + '%';
    if (lbl)  lbl.textContent = label;
  }
  function _indeterminatePct() { if (_progressTick < 90) _progressTick += 0.5; return _progressTick; }

  async function _pollTaskStatus(taskId, totalRows) {
    try {
      const resp = await fetch(`/device-monitoring/import/status/${taskId}/`, { credentials: 'same-origin' });
      const data = await resp.json().catch(() => null);
      if (!data) { _stopPolling(); showToast('Import polling failed', 'error'); return; }
      if (data.state === 'SUCCESS') {
        _stopPolling(); _setImportProgress(100, 'Done!');
        const sucEl = document.getElementById('import-success');
        if (sucEl) { sucEl.textContent = `✓ Import complete: ${data.created} created, ${data.updated} updated.`; sucEl.style.display = 'flex'; }
        const btn = document.getElementById('import-confirm-btn');
        if (btn) { btn.disabled = false; btn.textContent = 'Import'; }
        showToast(`✓ Import finished — ${data.created + data.updated} rows processed`, 'success');
        setTimeout(() => { closeImportModal(); window.location.reload(); }, 1800);
        return;
      }
      if (data.state === 'FAILURE') {
        _stopPolling();
        const errEl = document.getElementById('import-error');
        if (errEl) { errEl.textContent = 'Import failed: ' + (data.error || 'Unknown error'); errEl.style.display = 'flex'; }
        const btn = document.getElementById('import-confirm-btn');
        if (btn) { btn.disabled = false; btn.textContent = 'Import'; }
        return;
      }
      const pct = data.progress > 0 ? Math.min(95, data.progress) : _indeterminatePct();
      _setImportProgress(pct, data.message || `Processing ${totalRows} rows…`);
    } catch (err) {
      _stopPolling(); showToast('Network error: ' + err.message, 'error');
    }
  }

  async function confirmImport() {
    const fileInput = document.getElementById('import-file-input');
    const errEl     = document.getElementById('import-error');
    const sucEl     = document.getElementById('import-success');
    const btn       = document.getElementById('import-confirm-btn');
    if (!errEl || !sucEl || !btn || !fileInput?.files?.[0]) return;

    errEl.style.display = 'none'; sucEl.style.display = 'none';
    const formData = new FormData();
    formData.append('excel_file', fileInput.files[0]);
    formData.append('csrfmiddlewaretoken', getCsrf());

    btn.disabled = true;
    btn.innerHTML = '<svg style="width:14px;height:14px;animation:spin .7s linear infinite" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg> Uploading…';
    _progressTick = 10;
    _setImportProgress(5, 'Uploading file…');

    try {
      const resp = await fetch('/device-monitoring/import/', { method: 'POST', body: formData, credentials: 'same-origin' });
      const data = await resp.json();
      if (!resp.ok || !data.ok) throw new Error(data.error || `HTTP ${resp.status}`);
      if (data.done === true) {
        _setImportProgress(100, 'Done!');
        sucEl.textContent = `✓ Import complete: ${data.created} created, ${data.updated} updated.`;
        sucEl.style.display = 'flex';
        btn.disabled = false; btn.textContent = 'Import';
        showToast(`✓ Import finished — ${data.created + data.updated} rows`, 'success');
        setTimeout(() => { closeImportModal(); window.location.reload(); }, 1800);
        return;
      }
      if (!data.task_id) {
        _setImportProgress(100, 'Done!');
        sucEl.textContent = data.message || 'No data rows found.';
        sucEl.style.display = 'flex';
        btn.disabled = false; btn.textContent = 'Import';
        return;
      }
      _currentTask = data.task_id;
      _setImportProgress(15, `Queued ${data.total} rows…`);
      btn.innerHTML = '<svg style="width:14px;height:14px;animation:spin .7s linear infinite" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg> Processing…';
      _pollTaskStatus(data.task_id, data.total);
      _pollTimer = setInterval(() => _pollTaskStatus(data.task_id, data.total), 2000);
    } catch (err) {
      errEl.textContent = 'Error: ' + err.message; errEl.style.display = 'flex';
      btn.disabled = false; btn.textContent = 'Import';
      _setImportProgress(0, '');
    }
  }

  /* ==================== DRAG-TO-SCROLL ==================== */
  function initDragScroll(container) {
    if (!container) return;
    let isDragging = false, startX = 0, startY = 0, scrollLeft = 0, scrollTopStart = 0, hasDragged = false;
    container.addEventListener('mousedown', e => {
      if (e.button !== 0 || ['INPUT','TEXTAREA','SELECT','BUTTON','A','LABEL'].includes(e.target.tagName)) return;
      isDragging = true; hasDragged = false;
      startX = e.pageX - container.offsetLeft; startY = e.pageY - container.offsetTop;
      scrollLeft = container.scrollLeft; scrollTopStart = container.scrollTop;
      container.style.cursor = 'grabbing'; e.preventDefault();
    });
    document.addEventListener('mousemove', e => {
      if (!isDragging) return;
      const walkX = (e.pageX - container.offsetLeft) - startX;
      const walkY = (e.pageY - container.offsetTop) - startY;
      if (!hasDragged && (Math.abs(walkX) > 5 || Math.abs(walkY) > 5)) hasDragged = true;
      container.scrollLeft = scrollLeft - walkX;
      container.scrollTop  = scrollTopStart - walkY;
    });
    document.addEventListener('mouseup', () => { if (!isDragging) return; isDragging = false; container.style.cursor = ''; });
    container.addEventListener('click', e => { if (hasDragged) { e.stopPropagation(); e.preventDefault(); hasDragged = false; } }, true);
  }

  /* ==================== EVENT LISTENERS ==================== */
  function attachEventListeners() {
    document.addEventListener('focusin', e => {
      const tr = e.target.closest('tr[data-row-id]');
      focusedRowId = tr?.dataset?.rowId || null;
    });
    document.addEventListener('focusout', e => {
      const tr = e.target.closest('tr[data-row-id]');
      setTimeout(() => {
        const active = document.activeElement?.closest('tr[data-row-id]');
        if (!active) focusedRowId = null;
      }, 50);

      // Trigger auto-save (works for both new and existing rows)
      if (tr) scheduleAutoSave(tr);
    });

    // Checkboxes
    document.addEventListener('change', e => {
      const cb = e.target.closest('.dm-checkbox');
      if (cb?.type === 'checkbox') {
        syncCheck(cb);
        handleDmCheck(cb, cb.getAttribute('data-field'));
        return;
      }
      // borrower_type select
      const tr = e.target.closest('tr[data-row-id]');
      if (tr && e.target.matches('select[name="borrower_type"]')) {
        const rowId = tr.dataset.rowId;
        if (rowId && !rowId.startsWith('new_')) dirtyRows.add(rowId);
        const dataRow = allRows.find(r => String(r.id) === String(rowId));
        if (dataRow) dataRow.borrower_type = e.target.value;
        scheduleAutoSave(tr);
      }
    });

    // Text inputs / textareas
    document.addEventListener('input', e => {
      const tr = e.target.closest('tr[data-row-id]');
      if (!tr) return;
      const rowId = tr.dataset.rowId;
      if (rowId && !rowId.startsWith('new_')) dirtyRows.add(rowId);

      const dataRow = allRows.find(r => String(r.id) === String(rowId));
      if (dataRow) {
        if (e.target.name === 'box_number')          { dataRow.box_number = e.target.value; tr.dataset.box = e.target.value.toLowerCase(); }
        if (e.target.name === 'office_college')       { dataRow.office_college = e.target.value; tr.dataset.college = e.target.value.toLowerCase(); }
        if (e.target.name === 'accountable_person')   dataRow.accountable_person = e.target.value;
        if (e.target.name === 'accountable_officer')  { dataRow.accountable_officer = e.target.value; tr.dataset.officer = e.target.value.toLowerCase(); }
        if (e.target.name === 'device')               dataRow.device = e.target.value;
        if (e.target.name === 'serial_number')        dataRow.serial_number = e.target.value;
        if (e.target.name === 'assigned_mr')          { dataRow.assigned_mr = e.target.value; tr.dataset.mr = e.target.value; }
        if (e.target.name === 'ptr')                  { dataRow.ptr = e.target.value; tr.dataset.ptr = e.target.value; }
        if (e.target.name === 'remarks')              dataRow.remarks = e.target.value;
        if (e.target.name === 'issue')                dataRow.issue = e.target.value;
      }

      // For existing rows, schedule a debounced save on input.
      // For new rows, save is triggered by focusout, not by every keystroke.
      if (rowId && !rowId.startsWith('new_')) scheduleAutoSave(tr);
    });

    // Filters
    const si = document.getElementById('dm-search');
    if (si) si.addEventListener('input', debouncedFilter);
    ['dm-filter-college','dm-filter-borrower-type','dm-filter-officer',
     'dm-filter-mr','dm-filter-ptr','dm-filter-release','dm-filter-status'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', applyFilters);
    });
    document.getElementById('dm-clear-filters')?.addEventListener('click', () => {
      if (si) si.value = '';
      ['dm-filter-college','dm-filter-borrower-type','dm-filter-officer',
       'dm-filter-mr','dm-filter-ptr','dm-filter-release','dm-filter-status'].forEach(id => {
        const el = document.getElementById(id); if (el) el.selectedIndex = 0;
      });
      applyFilters();
    });

    document.getElementById('addDmRowBtn')?.addEventListener('click', addDmRow);
    document.getElementById('saveAllBtn')?.addEventListener('click', saveAllRows);

    document.addEventListener('click', e => {
      const db = e.target.closest('.dm-delete-row');
      if (db) { e.preventDefault(); deleteRow(db); }
    });

    document.getElementById('openImportModalBtn')?.addEventListener('click', openImportModal);
    document.getElementById('closeImportModalBtn')?.addEventListener('click', closeImportModal);
    document.getElementById('cancelImportBtn')?.addEventListener('click', closeImportModal);
    document.getElementById('import-confirm-btn')?.addEventListener('click', confirmImport);

    const fileInput = document.getElementById('import-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', function () {
        const errEl = document.getElementById('import-error');
        const prevEl = document.getElementById('import-preview');
        const btn    = document.getElementById('import-confirm-btn');
        if (errEl)  errEl.style.display  = 'none';
        if (prevEl) prevEl.style.display = 'none';
        if (btn)    btn.disabled = true;
        if (!this.files?.[0]) return;
        if (!this.files[0].name.match(/\.(xlsx|xls)$/i)) {
          if (errEl) { errEl.textContent = 'Please select a valid .xlsx or .xls file.'; errEl.style.display = 'flex'; }
          return;
        }
        const previewText = document.getElementById('import-preview-text');
        if (previewText) previewText.textContent = `File: ${this.files[0].name}\nSize: ${(this.files[0].size/1024).toFixed(1)} KB\nReady to import.`;
        if (prevEl) prevEl.style.display = 'block';
        if (btn)    btn.disabled = false;
      });
    }

    initDragScroll(document.querySelector('.table-container'));

    const exportBtn = document.querySelector('.export-btn');
    if (exportBtn) {
      exportBtn.addEventListener('click', async function(e) {
        e.preventDefault();
        const overlay = document.getElementById('invsys-loading-overlay');
        if (overlay) overlay.classList.add('is-active');
        try {
          const resp = await fetch(this.href);
          const data = await resp.json();
          if (data.ok && data.task_id) {
            pollExportTask(data.task_id);
          } else {
            if (overlay) overlay.classList.remove('is-active');
            alert('Export failed to start.');
          }
        } catch (err) {
          if (overlay) overlay.classList.remove('is-active');
          alert('Export failed to start.');
        }
      });
    }

    const tableContainer = document.querySelector('.table-container');
    if (tableContainer) attachReleaseEditListener(tableContainer);
  }

  /* ==================== SAVE ALL ROWS ==================== */
  async function saveAllRows() {
    // Save unsaved new rows first
    const newRowPromises = [];
    for (const [id, tr] of rowIdToElement) {
      if (id.startsWith('new_') && !_savingNew.has(id)) {
        newRowPromises.push(_saveNewRow(tr));
      }
    }

    // Then save dirty existing rows
    const existingPromises = [];
    for (const [id, tr] of rowIdToElement) {
      if (!id.startsWith('new_') && dirtyRows.has(id)) {
        existingPromises.push(saveRow(tr));
      }
    }

    await Promise.all([...newRowPromises, ...existingPromises]);
    const total = newRowPromises.length + existingPromises.length;
    if (total > 0) showToast(`Saved ${total} row(s)`, 'success');
    else showToast('Nothing to save', 'success');
  }

  /* ==================== INIT ==================== */
  document.addEventListener('DOMContentLoaded', () => {
    tbody = document.getElementById('dm-tbody');
    if (!tbody) return;

    const loadingRow = document.createElement('tr');
    loadingRow.id = 'dm-init-loading';
    loadingRow.innerHTML = `<td colspan="19" style="text-align:center;padding:48px;color:var(--muted)">
      <svg style="width:20px;height:20px;animation:spin .7s linear infinite;vertical-align:middle;margin-right:10px" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
        <path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/>
      </svg>
      Loading device rows…
    </td>`;
    tbody.appendChild(loadingRow);

    attachEventListeners();

    const dmUrl    = window.INVSYS_DM_AJAX || '/ajax/device-monitoring/';
    const embedded = typeof DM_ROWS !== 'undefined' && DM_ROWS.length > 0;

    const startRows = (rows) => {
      loadingRow.remove();
      initVirtualScroll(rows.map(normalizeDmRow));
    };

    if (embedded) {
      loadingRow.remove();
      initVirtualScroll(DM_ROWS.map(normalizeDmRow));
    } else {
      fetch(dmUrl, { credentials: 'same-origin', headers: { Accept: 'application/json' } })
        .then(r => { if (!r.ok) throw new Error('load failed'); return r.json(); })
        .then(data => startRows(data.rows || []))
        .catch(() => {
          loadingRow.innerHTML = '<td colspan="19" style="text-align:center;padding:48px;color:var(--muted)">Could not load device monitoring data.</td>';
          showToast('Could not load device monitoring', 'error');
        });
    }

    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined') {
      InvSysRT.connect('/ws/device-monitoring/', handleMessage, indicator);
    }
  });
})();