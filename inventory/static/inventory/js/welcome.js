/**
 * welcome.js – draggable dancing tablet diagnostic version
 */
(function () {
  'use strict';

  window.toggleBorrowerForms = function () {
    const selectedRole = document.querySelector('input[name="borrower_role"]:checked');
    const studentForm  = document.getElementById('student-form');
    const employeeForm = document.getElementById('employee-form');
    if (!selectedRole) return;
    const isStudent = selectedRole.value === 'student';
    if (studentForm)  studentForm.style.display  = isStudent ? 'block' : 'none';
    if (employeeForm) employeeForm.style.display  = isStudent ? 'none'  : 'block';

    const studentFields  = ['student_last_name','student_first_name','student_middle_initial','year_level','section','student_id','college','academic_year'];
    const employeeFields = ['employee_last_name','employee_first_name','employee_middle_initial','employee_id','office'];
    studentFields.forEach(f => {
      const el = document.querySelector(`[name="${f}"]`);
      if (el) el.disabled = !isStudent;
    });
    employeeFields.forEach(f => {
      const el = document.querySelector(`[name="${f}"]`);
      if (el) el.disabled = isStudent;
    });
  };

  window.closeSuccessModal = function () {
    const modal = document.getElementById('successModal');
    if (modal) modal.style.display = 'none';
  };

  /* ═══════════════════════════════════════════════════
     DRAG – directly on the tablet
  ═══════════════════════════════════════════════════ */
  function initDrag(tablet) {
    if (!tablet) {
      console.error('❌ initDrag: no element provided');
      return;
    }
    console.log('🔧 initDrag attached to', tablet);

    let offsetX, offsetY, startX, startY;
    let dragging = false;

    function onStart(e) {
      e.preventDefault();
      dragging = true;
      tablet.classList.add('dragging');
      console.log('✅ drag start');

      const rect = tablet.getBoundingClientRect();
      // Convert from right/bottom to left/top
      tablet.style.right = 'auto';
      tablet.style.bottom = 'auto';
      tablet.style.left = rect.left + 'px';
      tablet.style.top  = rect.top  + 'px';

      // For mouse
      startX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
      startY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;
      offsetX = startX - rect.left;
      offsetY = startY - rect.top;
    }

    function onMove(e) {
      if (!dragging) return;
      e.preventDefault();

      const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
      const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

      const newLeft = clientX - offsetX;
      const newTop  = clientY - offsetY;
      tablet.style.left = newLeft + 'px';
      tablet.style.top  = newTop + 'px';
    }

    function onEnd(e) {
      if (!dragging) return;
      dragging = false;
      tablet.classList.remove('dragging');
      console.log('✅ drag end');
    }

    // Mouse events
    tablet.addEventListener('mousedown', onStart);
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onEnd);

    // Touch events
    tablet.addEventListener('touchstart', onStart, { passive: false });
    document.addEventListener('touchmove', onMove, { passive: false });
    document.addEventListener('touchend', onEnd);
  }

  /* ── Init ────────────────────────────────────────────────── */
  document.addEventListener('DOMContentLoaded', function () {
    const tablet = document.querySelector('.dance-tablet');
    if (!tablet) {
      console.warn('❌ .dance-tablet not found');
      return;
    }
    // Check computed style
    const style = getComputedStyle(tablet);
    console.log('Computed pointer-events:', style.pointerEvents);
    console.log('Computed cursor:', style.cursor);

    initDrag(tablet);

    // Year warning (unchanged)
    const yearSelect = document.getElementById('year-level-select');
    const warningEl  = document.getElementById('fourth-year-warning');
    if (yearSelect && warningEl) {
      const check = () => {
        warningEl.style.display = /(4th|5th)/i.test(yearSelect.value) ? 'block' : 'none';
      };
      yearSelect.addEventListener('change', check);
      check();
    }

    const selRole = document.querySelector('input[name="borrower_role"]:checked');
    if (selRole) toggleBorrowerForms();

    document.addEventListener('click', function (e) {
      const modal = document.getElementById('successModal');
      if (modal && e.target === modal) closeSuccessModal();
    });
  });
})();