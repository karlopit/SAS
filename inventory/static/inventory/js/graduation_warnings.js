/**
 * graduation_warnings.js
 * Handles the client-side search filter for the Graduation Warnings table.
 */

(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    const input = document.getElementById('warn-search');
    if (!input) return;
    input.addEventListener('input', function () {
      const q = this.value.toLowerCase();
      document.querySelectorAll('#warn-tbody tr').forEach(row => {
        row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
      });
    });
  });
})();