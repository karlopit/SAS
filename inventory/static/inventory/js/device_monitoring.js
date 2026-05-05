/**
 * device_monitoring.js — Performance-optimized for 4k+ rows
 *
 * Key improvements vs original:
 * 1. VIRTUAL SCROLLING — only renders ~50 rows in DOM at a time, not all 4k+
 *    This is the #1 fix for the freeze/hang problem.
 * 2. Debounced filters — search/filter waits 200ms after typing stops
 * 3. Batched DOM saves — dirty row tracking prevents unnecessary re-renders
 * 4. Import progress polling unchanged (already async/Celery-based)
 */

(function () {
  'use strict';

  /* ==================== CONSTANTS ==================== */
  const ROW_HEIGHT     = 64;   // px — must match actual rendered row height
  const OVERSCAN       = 10;   // extra rows above/below visible area
  const VISIBLE_BUFFER = 15;   // rows kept outside viewport for smooth scroll

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
    }, 4000);
  }

  /* ==================== DEBOUNCE ==================== */
  function debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  }

  /** Align WebSocket / AJAX row shape with table (date label vs raw). */
  function normalizeDmRow(r) {
    const o = { ...r };
    o.date_returned_display = o.date_returned_display ?? o.date_returned ?? '—';
    return o;
  }

  /* ==================== STATE ==================== */
  // allRows = full dataset from DM_ROWS (never mutated)
  // filteredRows = after filters/search (rebuilt on filter change)
  // dirtyRows = Set of row IDs with unsaved user edits
  let allRows      = [];
  let filteredRows = [];
  const dirtyRows  = new Set();

  // Virtual scroll state
  let scrollTop      = 0;
  let containerHeight = 600;
  let startIdx = 0, endIdx = 0;

  // DOM references (set after DOMContentLoaded)
  let tbody, scrollContainer, topSpacer, bottomSpacer;

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

  /* ==================== BUILD ROW HTML ==================== */
  function _esc(str) {
    return String(str || '')
      .replace(/&/g, '&amp;').replace(/"/g, '&quot;')
      .replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function buildRowHtml(row) {
    const releaseClass = row.release_status === 'Released' ? 'badge-released'
                       : row.release_status === 'Returned' ? 'badge-returned-dm'
                       : 'badge-none';

    return `<input type="hidden" name="row_id" value="${row.id}"/>
      <td style="text-align:center"><input type="text" name="box_number" value="${_esc(row.box_number)}" class="form-control dm-box-input" placeholder="Box #" style="width:80px;text-align:center;margin:0 auto"/></td>
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
      <td style="text-align:center"><input type="text" name="assigned_mr" value="${_esc(row.assigned_mr)}" class="form-control dm-mr-input" placeholder="M.R. #" style="width:110px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="accountable_officer" value="${_esc(row.accountable_officer)}" class="form-control dm-officer-input" placeholder="Officer name" style="width:130px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="device" value="${_esc(row.device)}" class="form-control dm-device-input" style="width:90px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="serviceable" value="${row.serviceable ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="serviceable" ${row.serviceable ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="non_serviceable" value="${row.non_serviceable ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="non_serviceable" ${row.non_serviceable ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="sealed" value="${row.sealed ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="sealed" ${row.sealed ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="missing" value="${row.missing ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="missing" ${row.missing ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="incomplete" value="${row.incomplete ? 'on' : 'off'}"/><input type="checkbox" class="dm-checkbox" data-field="incomplete" ${row.incomplete ? 'checked' : ''} style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="ptr" value="${_esc(row.ptr)}" class="form-control dm-ptr-input" placeholder="PTR #" style="width:100px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><span class="release-status-badge ${releaseClass}">${_esc(row.release_status)}</span></td>
      <td style="text-align:center;color:var(--muted);font-size:12px" class="dm-date-returned">${_esc(row.date_returned_display || '—')}</td>
      <td style="text-align:center"><textarea name="remarks" class="form-control dm-remarks-input" rows="2" placeholder="Remarks…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto">${_esc(row.remarks)}</textarea></td>
      <td style="text-align:center"><textarea name="issue" class="form-control dm-issue-input" rows="2" placeholder="Issue…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto">${_esc(row.issue)}</textarea></td>
      <td style="text-align:center;white-space:nowrap">
        <button type="submit" class="btn btn-primary btn-sm">✓ Save</button>
        <button type="button" class="btn btn-danger btn-sm dm-delete-row" style="margin-left:4px">✕</button>
      </td>`;
  }

  /* ==================== VIRTUAL SCROLL ENGINE ==================== */
  // rowIdToElement: maps rowId → tr (for dirty/focused rows kept in DOM)
  const rowIdToElement = new Map();
  // focusedRowId: row the user is currently editing — never remove from DOM
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

    // Update spacers so scrollbar is accurate
    topSpacer.style.height    = (newStart * ROW_HEIGHT) + 'px';
    bottomSpacer.style.height = (Math.max(0, filteredRows.length - newEnd) * ROW_HEIGHT) + 'px';

    // Build set of IDs that should be rendered
    const neededIds = new Set();
    for (let i = newStart; i < newEnd; i++) {
      neededIds.add(String(filteredRows[i].id));
    }
    // Always keep focused/dirty rows in DOM
    if (focusedRowId) neededIds.add(focusedRowId);
    dirtyRows.forEach(id => neededIds.add(id));

    // Remove rows no longer needed
    rowIdToElement.forEach((tr, id) => {
      if (!neededIds.has(id)) {
        // Harvest edits back to allRows data before removing
        harvestRowEdits(tr, id);
        tr.remove();
        rowIdToElement.delete(id);
      }
    });

    // Build fragment for new rows (in order)
    const fragment = document.createDocumentFragment();
    const insertedIds = new Set();

    for (let i = newStart; i < newEnd; i++) {
      const row = filteredRows[i];
      const strId = String(row.id);
      if (rowIdToElement.has(strId)) {
        insertedIds.add(strId);
        continue; // already in DOM
      }
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
      insertedIds.add(strId);
      fragment.appendChild(tr);
    }

    // Insert fragment before bottomSpacer
    tbody.insertBefore(fragment, bottomSpacer);

    startIdx = newStart;
    endIdx   = newEnd;

    // Update filter count
    const vc = document.getElementById('dm-visible-count');
    const tc = document.getElementById('dm-total-count');
    if (vc) vc.textContent = filteredRows.length;
    if (tc) tc.textContent = allRows.length;
  }

  function harvestRowEdits(tr, id) {
    // Write DOM values back to allRows so they survive a virtual scroll cycle
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
    const cbState = (name) => tr.querySelector(`input[type=hidden][name="${name}"]`)?.value === 'on';
    dataRow.serviceable     = cbState('serviceable');
    dataRow.non_serviceable = cbState('non_serviceable');
    dataRow.sealed          = cbState('sealed');
    dataRow.missing         = cbState('missing');
    dataRow.incomplete      = cbState('incomplete');
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

    // Flush rendered rows so virtual scroll rebuilds with filtered set
    rowIdToElement.forEach(tr => { harvestRowEdits(tr, tr.dataset.rowId); tr.remove(); });
    rowIdToElement.clear();

    // Reset scroll and re-render
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

    // Create spacer rows
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

    // Resize observer for responsive containers
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
  function addDmRow() {
    const newId = 'new_' + Date.now();
    const newRow = {
      id: newId, box_number: '', serial_number: '', office_college: '',
      accountable_person: '', borrower_type: '', assigned_mr: '',
      accountable_officer: '', device: 'Tablet', serviceable: false,
      non_serviceable: false, sealed: false, missing: false, incomplete: false,
      ptr: '', remarks: '', issue: '', release_status: '—',
      date_returned_display: '—',
    };

    // Add to data arrays
    allRows.push(newRow);
    filteredRows.push(newRow);

    // Build and append TR directly — new rows always visible at bottom
    const tr = document.createElement('tr');
    tr.dataset.rowId = newId;
    tr.innerHTML = buildRowHtml(newRow);
    applyLockState(tr);
    tbody.insertBefore(tr, bottomSpacer);
    rowIdToElement.set(newId, tr);

    // Scroll to bottom
    if (scrollContainer) scrollContainer.scrollTop = scrollContainer.scrollHeight;
    tr.querySelector('input[name="box_number"]')?.focus();

    const tc = document.getElementById('dm-total-count');
    if (tc) tc.textContent = allRows.length;
  }

  /* ==================== SAVE ALL ROWS (JSON) ==================== */
  async function saveAllRows() {
    const form = document.getElementById('dm-form');
    if (!form) return;

    const btn = document.getElementById('saveAllBtn');
    const originalHTML = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<svg style="width:14px;height:14px;animation:spin .7s linear infinite" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg> Saving…';
    }

    // Harvest all in-DOM edits back to allRows before building payload
    rowIdToElement.forEach((tr, id) => harvestRowEdits(tr, id));

    // Build payload from in-memory data (includes all rows, not just visible ones)
    const rowsData = allRows.map(row => ({
      row_id:              String(row.id),
      box_number:          row.box_number          || '',
      serial_number:       row.serial_number       || '',
      office_college:      row.office_college      || '',
      accountable_person:  row.accountable_person  || '',
      borrower_type:       row.borrower_type       || '',
      assigned_mr:         row.assigned_mr         || '',
      accountable_officer: row.accountable_officer || '',
      device:              row.device              || 'Tablet',
      serviceable:         row.serviceable     ? 'on' : 'off',
      non_serviceable:     row.non_serviceable ? 'on' : 'off',
      sealed:              row.sealed          ? 'on' : 'off',
      missing:             row.missing         ? 'on' : 'off',
      incomplete:          row.incomplete      ? 'on' : 'off',
      ptr:                 row.ptr                 || '',
      remarks:             row.remarks             || '',
      issue:               row.issue               || '',
    }));

    try {
      const resp = await fetch(form.action, {
        method:  'POST',
        headers: {
          'Content-Type':     'application/json',
          'X-CSRFToken':      document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '',
          'X-Requested-With': 'XMLHttpRequest',
        },
        body:        JSON.stringify({ rows: rowsData, save_all: true }),
        credentials: 'same-origin',
      });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const result = await resp.json();
      if (result.ok) {
        dirtyRows.clear();
        showToast(`✓ All rows saved (${result.saved} record${result.saved !== 1 ? 's' : ''})`, 'success');
      } else {
        showToast(result.errors?.length ? `Saved with errors: ${result.errors.slice(0, 2).join('; ')}` : 'Save failed', 'error');
      }
    } catch (err) {
      showToast('Network error: ' + err.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = originalHTML; }
    }
  }

  /* ==================== DELETE ROW ==================== */
  function deleteRow(btn) {
    const tr    = btn.closest('tr');
    const rowId = tr?.dataset?.rowId;
    if (!rowId) return;

    if (!rowId.startsWith('new_')) {
      if (!confirm('Delete this row?')) return;
      const form  = document.createElement('form');
      form.method = 'post';
      form.action = `/device-monitoring/${rowId}/delete/`;
      const csrf  = document.createElement('input');
      csrf.type = 'hidden'; csrf.name = 'csrfmiddlewaretoken';
      csrf.value = document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '';
      form.appendChild(csrf);
      document.body.appendChild(form);
      form.submit();
    } else {
      allRows      = allRows.filter(r => String(r.id) !== rowId);
      filteredRows = filteredRows.filter(r => String(r.id) !== rowId);
      rowIdToElement.delete(rowId);
      tr.remove();
      const tc = document.getElementById('dm-total-count');
      if (tc) tc.textContent = allRows.length;
    }
  }

  /* ==================== WEBSOCKET REALTIME ==================== */
  function handleMessage(data) {
    if (data.type !== 'device_monitoring.update') return;

    // Update in-memory data (don't touch dirty/focused rows in DOM)
    const incoming = new Map(data.rows.map(r => [String(r.id), r]));

    // Update allRows in place
    for (let i = 0; i < allRows.length; i++) {
      const id = String(allRows[i].id);
      if (incoming.has(id) && !dirtyRows.has(id)) {
        const inc = normalizeDmRow(incoming.get(id));
        allRows[i] = { ...allRows[i], ...inc };
        const tr = rowIdToElement.get(id);
        if (tr && id !== focusedRowId && !dirtyRows.has(id)) {
          const row = allRows[i];
          const badge = tr.querySelector('.release-status-badge');
          if (badge) {
            const cls = row.release_status === 'Released' ? 'badge-released'
                       : row.release_status === 'Returned' ? 'badge-returned-dm'
                       : 'badge-none';
            badge.className = `release-status-badge ${cls}`;
            badge.textContent = row.release_status || '—';
          }
          const dateTd = tr.querySelector('.dm-date-returned');
          if (dateTd) dateTd.textContent = row.date_returned_display || '—';
        }
      }
    }

    if (typeof data.pending_count === 'number')
      window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.pending_count }));
    if (typeof data.graduation_warning_count === 'number')
      window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));
  }

  /* ==================== IMPORT MODAL ==================== */
  function _getCsrf() { return document.cookie.match(/csrftoken=([^;]+)/)?.[1] || ''; }

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
        if (sucEl) {
          sucEl.textContent = `✓ Import complete: ${data.created} created, ${data.updated} updated.`;
          sucEl.style.display = 'flex';
        }
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
    formData.append('csrfmiddlewaretoken', _getCsrf());

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
    // Focus/blur tracking for virtual scroll protection
    document.addEventListener('focusin', e => {
      const tr = e.target.closest('tr[data-row-id]');
      focusedRowId = tr?.dataset?.rowId || null;
    });
    document.addEventListener('focusout', () => {
      setTimeout(() => {
        const active = document.activeElement?.closest('tr[data-row-id]');
        if (!active) focusedRowId = null;
      }, 50);
    });

    // Checkboxes
    document.addEventListener('change', e => {
      const cb = e.target.closest('.dm-checkbox');
      if (cb?.type === 'checkbox') {
        syncCheck(cb);
        handleDmCheck(cb, cb.getAttribute('data-field'));
      }
    });

    // Input changes — update in-memory data + mark dirty
    document.addEventListener('input', e => {
      const row = e.target.closest('tr[data-row-id]');
      if (!row) return;
      const rowId = row.dataset.rowId;
      if (rowId && !rowId.startsWith('new_')) dirtyRows.add(rowId);

      // Update in-memory row immediately
      const dataRow = allRows.find(r => String(r.id) === String(rowId));
      if (dataRow) {
        if (e.target.name === 'box_number')          { dataRow.box_number = e.target.value; row.dataset.box = e.target.value.toLowerCase(); }
        if (e.target.name === 'office_college')       { dataRow.office_college = e.target.value; row.dataset.college = e.target.value.toLowerCase(); }
        if (e.target.name === 'accountable_person')   dataRow.accountable_person = e.target.value;
        if (e.target.name === 'borrower_type')        { dataRow.borrower_type = e.target.value; row.dataset.borrowerType = e.target.value; }
        if (e.target.name === 'accountable_officer')  { dataRow.accountable_officer = e.target.value; row.dataset.officer = e.target.value.toLowerCase(); }
        if (e.target.name === 'device')               dataRow.device = e.target.value;
        if (e.target.name === 'serial_number')        dataRow.serial_number = e.target.value;
        if (e.target.name === 'assigned_mr')          { dataRow.assigned_mr = e.target.value; row.dataset.mr = e.target.value; }
        if (e.target.name === 'ptr')                  { dataRow.ptr = e.target.value; row.dataset.ptr = e.target.value; }
        if (e.target.name === 'remarks')              dataRow.remarks = e.target.value;
        if (e.target.name === 'issue')                dataRow.issue = e.target.value;
      }
    });

    // Filters — debounced
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

    const dmUrl = window.INVSYS_DM_AJAX || '/ajax/device-monitoring/';
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
        .then((r) => {
          if (!r.ok) throw new Error('load failed');
          return r.json();
        })
        .then((data) => startRows(data.rows || []))
        .catch(() => {
          loadingRow.innerHTML = '<td colspan="19" style="text-align:center;padding:48px;color:var(--muted)">Could not load device monitoring data.</td>';
          showToast('Could not load device monitoring', 'error');
        });
    }

    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined') {
      InvSysRT.connect('/ws/device-monitoring/', handleMessage, indicator);
    }
    
    // ── Export Excel button (Celery async) ──
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
  });
})();