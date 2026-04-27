/**
 * staff_confirm_borrow.js
 * Handles quantity hint update and real-time serial/box number counting
 * on the staff confirm-borrow form.
 */

(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    const itemSelect            = document.getElementById('id_item');
    const quantityInput         = document.getElementById('id_quantity_borrowed');
    const hintSpan              = document.getElementById('qty-hint');
    const serialNumbersTextarea = document.getElementById('id_serial_numbers');
    const boxNumbersTextarea    = document.getElementById('id_box_numbers');
    const serialCountHint       = document.getElementById('serial-count-hint');
    const boxCountHint          = document.getElementById('box-count-hint');

    /* ── Quantity / availability hint ──────────────────────────────────── */
    function updateQuantityHint() {
      if (!itemSelect) return;
      const opt    = itemSelect.options[itemSelect.selectedIndex];
      const maxQty = opt.getAttribute('data-available');
      if (maxQty && maxQty !== 'None') {
        if (quantityInput) quantityInput.max = maxQty;
        if (hintSpan) hintSpan.textContent = `Maximum available: ${maxQty}`;
        if (quantityInput && parseInt(quantityInput.value) > parseInt(maxQty)) {
          quantityInput.value = maxQty;
        }
      } else {
        if (hintSpan) hintSpan.textContent = 'Select an item to see availability.';
      }
    }

    /* ── Line counter ──────────────────────────────────────────────────── */
    function getLineCount(textarea) {
      if (!textarea) return 0;
      return textarea.value.split('\n').filter(l => l.trim().length > 0).length;
    }

    function updateSerialCount() {
      if (!serialNumbersTextarea || !serialCountHint) return;
      const count      = getLineCount(serialNumbersTextarea);
      const required   = parseInt(quantityInput?.value) || 0;
      if (count === 0) {
        serialCountHint.innerHTML = '';
      } else if (count === required) {
        serialCountHint.innerHTML = `✓ ${count} of ${required} serial numbers entered. Ready to submit!`;
        serialCountHint.style.color = '#00e5a0';
      } else {
        serialCountHint.innerHTML = `⚠️ ${count} of ${required} serial numbers entered. Need ${required - count} more.`;
        serialCountHint.style.color = '#ffc107';
      }
    }

    function updateBoxCount() {
      if (!boxNumbersTextarea || !boxCountHint) return;
      const count    = getLineCount(boxNumbersTextarea);
      const required = parseInt(quantityInput?.value) || 0;
      if (count === 0) {
        boxCountHint.innerHTML = '';
      } else if (count === required) {
        boxCountHint.innerHTML = `✓ ${count} of ${required} box numbers entered. Ready to submit!`;
        boxCountHint.style.color = '#00e5a0';
      } else {
        boxCountHint.innerHTML = `⚠️ ${count} of ${required} box numbers entered. Need ${required - count} more.`;
        boxCountHint.style.color = '#ffc107';
      }
    }

    /* ── Event listeners ───────────────────────────────────────────────── */
    if (itemSelect) {
      itemSelect.addEventListener('change', updateQuantityHint);
      updateQuantityHint();
    }
    if (serialNumbersTextarea) serialNumbersTextarea.addEventListener('input', updateSerialCount);
    if (boxNumbersTextarea)    boxNumbersTextarea.addEventListener('input', updateBoxCount);
    if (quantityInput) {
      quantityInput.addEventListener('change', () => {
        updateSerialCount();
        updateBoxCount();
      });
    }

    updateSerialCount();
    updateBoxCount();
  });
})();