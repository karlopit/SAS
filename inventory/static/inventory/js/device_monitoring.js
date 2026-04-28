/**
 * device_monitoring.js
 * Handles:
 *  - Checkbox mutual‑exclusion logic
 *  - Filter dropdowns and search
 *  - Add/delete rows (PTR + Assigned M.R. included)
 *  - Excel import modal and file upload
 *  - Real‑time WebSocket updates
 */

(function () {
  'use strict';

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

  window.updateRowDataAttrs = function (cb) {
    const row = cb.closest('tr[data-row-id]');
    if (!row) return;
    const c = getChecksInRow(row);
    row.dataset.serviceable    = c.serviceable.cb?.checked    ? '1' : '0';
    row.dataset.nonServiceable = c.non_serviceable.cb?.checked ? '1' : '0';
    row.dataset.sealed         = c.sealed.cb?.checked         ? '1' : '0';
    row.dataset.missing        = c.missing.cb?.checked        ? '1' : '0';
    row.dataset.incomplete     = c.incomplete.cb?.checked     ? '1' : '0';
    applyDmFilters();
  };

  /* ==================== FILTER DROPDOWNS ==================== */
  function populateFilterDropdowns() {
    const collegeSelect = document.getElementById('dm-filter-college');
    const personSelect  = document.getElementById('dm-filter-person');
    const officerSelect = document.getElementById('dm-filter-officer');

    if (!collegeSelect || !personSelect || !officerSelect) return;

    while (collegeSelect.options.length > 1) collegeSelect.remove(1);
    while (personSelect.options.length  > 1) personSelect.remove(1);
    while (officerSelect.options.length > 1) officerSelect.remove(1);

    const colleges = new Set(), persons = new Set(), officers = new Set();

    document.querySelectorAll('#dm-tbody tr[data-row-id]').forEach(row => {
      const college = row.querySelector('input[name="office_college"]')?.value;
      const person  = row.querySelector('input[name="accountable_person"]')?.value;
      const officer = row.querySelector('input[name="accountable_officer"]')?.value;
      if (college && college.trim()) colleges.add(college);
      if (person  && person.trim())  persons.add(person);
      if (officer && officer.trim()) officers.add(officer);
    });

    const addOpts = (sel, set) => [...set].sort().forEach(v => {
      const o = document.createElement('option'); o.value = o.textContent = v; sel.appendChild(o);
    });
    addOpts(collegeSelect, colleges);
    addOpts(personSelect,  persons);
    addOpts(officerSelect, officers);
  }

  function applyDmFilters() {
    const search       = document.getElementById('dm-search').value.toLowerCase().trim();
    const college      = document.getElementById('dm-filter-college').value;
    const person       = document.getElementById('dm-filter-person').value;
    const borrowerType = document.getElementById('dm-filter-borrower-type').value;
    const officer      = document.getElementById('dm-filter-officer').value;
    const releaseF     = document.getElementById('dm-filter-release').value;
    const status       = document.getElementById('dm-filter-status').value;

    const rows = document.querySelectorAll('#dm-tbody tr[data-row-id]');
    let visible = 0;

    rows.forEach(row => {
      const rowText       = row.textContent.toLowerCase();
      const matchSearch   = !search      || rowText.includes(search);
      const matchCollege  = !college      || (row.dataset.college     || '').toLowerCase() === college.toLowerCase();
      const matchPerson   = !person       || (row.dataset.person      || '').toLowerCase() === person.toLowerCase();
      const matchBT       = !borrowerType || (row.dataset.borrowerType|| '').toLowerCase() === borrowerType.toLowerCase();
      const matchOfficer  = !officer      || (row.dataset.officer     || '').toLowerCase() === officer.toLowerCase();
      const matchRelease  = !releaseF     || (row.dataset.release     || '')               === releaseF;
      let   matchStatus   = true;
      if (status) {
        const attrKey = 'data-' + status.replace(/_/g, '-');
        matchStatus = row.getAttribute(attrKey) === '1';
      }

      const show = matchSearch && matchCollege && matchPerson && matchBT && matchOfficer && matchRelease && matchStatus;
      row.style.display = show ? '' : 'none';
      if (show) visible++;
    });

    const hasFilter = search || college || person || borrowerType || officer || releaseF || status;
    const statusLbl = document.getElementById('dm-filter-status-label');
    const clearBtn  = document.getElementById('dm-clear-filters');
    if (statusLbl) statusLbl.style.display = hasFilter ? 'inline-flex' : 'none';
    if (clearBtn)  clearBtn.style.display  = hasFilter ? 'inline-flex' : 'none';
    const vc = document.getElementById('dm-visible-count');
    const tc = document.getElementById('dm-total-count');
    if (vc) vc.textContent = visible;
    if (tc) tc.textContent = rows.length;
  }

  window.clearDmFilters = function () {
    const si = document.getElementById('dm-search');
    if (si) si.value = '';
    ['dm-filter-college', 'dm-filter-person', 'dm-filter-borrower-type',
     'dm-filter-officer', 'dm-filter-release', 'dm-filter-status'].forEach(id => {
      const el = document.getElementById(id); if (el) el.selectedIndex = 0;
    });
    populateFilterDropdowns();
    applyDmFilters();
  };

  /* ==================== ADD ROW (includes PTR + Assigned M.R.) ==================== */
  window.addDmRow = function () {
    const tbody = document.getElementById('dm-tbody');
    const tr    = document.createElement('tr');
    const newId = 'new_' + Date.now();
    const attrs = {
      'data-row-id': newId, 'data-box': '', 'data-college': '',
      'data-person': '', 'data-borrower-type': '', 'data-officer': '',
      'data-device': '', 'data-serial': '', 'data-release': '—',
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
      <td style="text-align:center"><input type="hidden" name="serviceable" value="off"/><input type="checkbox" onchange="syncCheck(this);handleDmCheck(this,'serviceable');updateRowDataAttrs(this)" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="non_serviceable" value="off"/><input type="checkbox" onchange="syncCheck(this);handleDmCheck(this,'non_serviceable');updateRowDataAttrs(this)" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="sealed" value="off"/><input type="checkbox" onchange="syncCheck(this);handleDmCheck(this,'sealed');updateRowDataAttrs(this)" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="missing" value="off"/><input type="checkbox" onchange="syncCheck(this);handleDmCheck(this,'missing');updateRowDataAttrs(this)" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="hidden" name="incomplete" value="off"/><input type="checkbox" onchange="syncCheck(this);handleDmCheck(this,'incomplete');updateRowDataAttrs(this)" style="margin:0 auto"/></td>
      <td style="text-align:center"><input type="text" name="ptr" class="form-control dm-ptr-input" placeholder="PTR #" style="width:100px;text-align:center;margin:0 auto"/></td>
      <td style="text-align:center"><span class="release-status-badge badge-none">—</span></td>
      <td style="text-align:center;color:var(--muted);font-size:12px" class="dm-date-returned">—</td>
      <td style="text-align:center"><textarea name="remarks" class="form-control dm-remarks-input" rows="2" placeholder="Remarks…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto"></textarea></td>
      <td style="text-align:center"><textarea name="issue" class="form-control dm-issue-input" rows="2" placeholder="Issue…" style="width:155px;font-size:12px;resize:vertical;margin:0 auto"></textarea></td>
      <td style="text-align:center;white-space:nowrap">
        <button type="submit" class="btn btn-primary btn-sm">✓ Save</button>
        <button type="button" class="btn btn-danger btn-sm" style="margin-left:4px"
          onclick="this.closest('tr').remove(); populateFilterDropdowns(); applyDmFilters();">✕</button>
      </td>
    `;
    tbody.appendChild(tr);
    tr.querySelector('input[name="box_number"]').focus();
    applyLockState(tr);
    populateFilterDropdowns();
    applyDmFilters();
  };

  /* ==================== IMPORT EXCEL MODAL ==================== */
  window.openImportModal = function () {
    const modal = document.getElementById('importModal');
    if (!modal) return;
    modal.style.display = 'flex';
    // reset state
    const fileInput = document.getElementById('import-file-input');
    if (fileInput) fileInput.value = '';
    const preview = document.getElementById('import-preview');
    const errorEl = document.getElementById('import-error');
    const successEl = document.getElementById('import-success');
    const confirmBtn = document.getElementById('import-confirm-btn');
    if (preview) preview.style.display = 'none';
    if (errorEl) errorEl.style.display = 'none';
    if (successEl) successEl.style.display = 'none';
    if (confirmBtn) confirmBtn.disabled = true;
  };

  window.closeImportModal = function (e) {
    const modal = document.getElementById('importModal');
    if (!modal) return;
    if (e && e.target !== modal) return;
    modal.style.display = 'none';
  };

  window.confirmImport = async function () {
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
    // retrieve CSRF token from cookie
    const csrfToken = document.cookie.match(/csrftoken=([^;]+)/)?.[1] || '';
    formData.append('csrfmiddlewaretoken', csrfToken);

    btn.disabled = true;
    const originalText = btn.innerHTML;
    btn.textContent = 'Importing…';

    try {
      const resp = await fetch('/device-monitoring/import/', {
        method: 'POST',
        body: formData,
        credentials: 'same-origin',
      });
      const data = await resp.json();

      if (!data.ok) throw new Error(data.error || 'Import failed');

      sucEl.textContent = `✓ Import complete: ${data.created} row(s) created, ${data.updated} row(s) updated.`;
      sucEl.style.display = 'flex';

      // reload table after short delay
      setTimeout(() => { window.location.reload(); }, 1500);
    } catch (err) {
      errEl.textContent = 'Error: ' + err.message;
      errEl.style.display = 'flex';
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
  };

  /* ==================== WEBSOCKET REAL‑TIME UPDATES ==================== */
  const dirtyRows = new Set();

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
      if (id && !id.startsWith('new_') && !incomingIds.has(id) && !dirtyRows.has(id)) {
        tr.remove();
      }
    });

    data.rows.forEach(row => {
      const strId    = String(row.id);
      const existing = tbody.querySelector(`tr[data-row-id="${strId}"]`);

      if (existing) {
        if (dirtyRows.has(strId)) return;

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

        existing.dataset.box          = (row.box_number || '').toLowerCase();
        existing.dataset.college      = (row.office_college || '').toLowerCase();
        existing.dataset.person       = (row.accountable_person || '').toLowerCase();
        existing.dataset.borrowerType = (row.borrower_type || '').toLowerCase();
        existing.dataset.officer      = (row.accountable_officer || '').toLowerCase();
        existing.dataset.device       = (row.device || '').toLowerCase();
        existing.dataset.serial       = (row.serial_number || '').toLowerCase();
        existing.dataset.release      = row.release_status || '—';
        existing.dataset.serviceable    = row.serviceable     ? '1' : '0';
        existing.dataset.nonServiceable = row.non_serviceable ? '1' : '0';
        existing.dataset.sealed         = row.sealed          ? '1' : '0';
        existing.dataset.missing        = row.missing         ? '1' : '0';
        existing.dataset.incomplete     = row.incomplete      ? '1' : '0';

        const releaseTd = existing.querySelector('.release-status-badge');
        if (releaseTd) releaseTd.outerHTML = releaseBadgeHtml(row.release_status);

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
      }
    });

    populateFilterDropdowns();
    applyDmFilters();

    if (data.pending_count !== undefined) {
      window.dispatchEvent(new CustomEvent('invsys:pending_count', { detail: data.pending_count }));
    }
    if (data.graduation_warning_count !== undefined) {
      window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count }));
    }
  }

  /* ==================== INITIALISATION ==================== */
  document.addEventListener('DOMContentLoaded', () => {
    // Apply checkbox lock states to existing rows
    document.querySelectorAll('#dm-tbody tr[data-row-id]').forEach(row => applyLockState(row));
    populateFilterDropdowns();

    // Mark rows as dirty when user edits
    const dmForm = document.getElementById('dm-form');
    if (dmForm) {
      dmForm.addEventListener('input', e => {
        const row = e.target.closest('tr[data-row-id]');
        if (row && row.dataset.rowId && !row.dataset.rowId.startsWith('new_')) {
          dirtyRows.add(row.dataset.rowId);
        }
      });
    }

    // Update dataset attributes on user input (for filtering)
    document.addEventListener('input', e => {
      const row = e.target.closest('tr[data-row-id]');
      if (!row) return;
      if (e.target.matches('.dm-box-input'))            { row.dataset.box         = e.target.value.toLowerCase(); applyDmFilters(); }
      if (e.target.matches('.dm-college-input'))        { row.dataset.college      = e.target.value.toLowerCase(); populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-person-input'))         { row.dataset.person       = e.target.value.toLowerCase(); populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-borrower-type-select')) { row.dataset.borrowerType = e.target.value.toLowerCase(); applyDmFilters(); }
      if (e.target.matches('.dm-officer-input'))        { row.dataset.officer      = e.target.value.toLowerCase(); populateFilterDropdowns(); applyDmFilters(); }
      if (e.target.matches('.dm-device-input'))         { row.dataset.device       = e.target.value.toLowerCase(); applyDmFilters(); }
      if (e.target.matches('.dm-serial-input'))         { row.dataset.serial       = e.target.value.toLowerCase(); applyDmFilters(); }
      // ptr and assigned_mr don't need dataset tracking (not filterable)
    });

    // Attach filter event listeners
    const searchInput = document.getElementById('dm-search');
    if (searchInput) searchInput.addEventListener('input', applyDmFilters);

    const filterIds = ['dm-filter-college', 'dm-filter-person', 'dm-filter-borrower-type',
                       'dm-filter-officer', 'dm-filter-release', 'dm-filter-status'];
    filterIds.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', applyDmFilters);
    });

    // Import modal: file input preview
    const fileInput = document.getElementById('import-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', function () {
        const file = this.files[0];
        const errEl = document.getElementById('import-error');
        const prevEl = document.getElementById('import-preview');
        const btn = document.getElementById('import-confirm-btn');
        if (errEl) errEl.style.display = 'none';
        if (prevEl) prevEl.style.display = 'none';
        if (btn) btn.disabled = true;

        if (!file) return;
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
          if (errEl) {
            errEl.textContent = 'Please select a valid .xlsx or .xls file.';
            errEl.style.display = 'flex';
          }
          return;
        }

        const rowCountSpan = document.getElementById('import-row-count');
        const previewText = document.getElementById('import-preview-text');
        if (rowCountSpan) rowCountSpan.textContent = '—';
        if (previewText) previewText.textContent = `File: ${file.name}\nSize: ${(file.size / 1024).toFixed(1)} KB\nReady to import.`;
        if (prevEl) prevEl.style.display = 'block';
        if (btn) btn.disabled = false;
      });
    }

    // WebSocket connection (if global InvSysRT exists)
    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined' && InvSysRT.connect) {
      InvSysRT.connect('/ws/device-monitoring/', handleMessage, indicator);
    }
  });
})();