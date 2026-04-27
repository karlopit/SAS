/**
 * welcome.js
 * Handles the public borrow-request form:
 *  - Borrower type toggle (student / employee)
 *  - 4th/5th year graduation warning display
 *  - Success modal close behaviour
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

  document.addEventListener('DOMContentLoaded', function () {
    /* Year-level graduation warning */
    const yearSelect = document.getElementById('year-level-select');
    const warningEl  = document.getElementById('fourth-year-warning');

    function checkYearWarning() {
      if (!yearSelect || !warningEl) return;
      const val = yearSelect.value.toLowerCase();
      warningEl.style.display = (val.includes('4th') || val.includes('5th')) ? 'block' : 'none';
    }

    if (yearSelect) yearSelect.addEventListener('change', checkYearWarning);
    checkYearWarning();

    /* Re-show correct sub-form if radio is already selected (after POST error) */
    const selectedRole = document.querySelector('input[name="borrower_role"]:checked');
    if (selectedRole) toggleBorrowerForms();

    /* Click outside success modal to close */
    document.addEventListener('click', function (e) {
      const modal = document.getElementById('successModal');
      if (modal && e.target === modal) closeSuccessModal();
    });
  });
})();