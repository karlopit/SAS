/**
 * index_admin.js
 * Handles the admin dashboard's edit-item modal and the sidebar pending-count
 * badge update via WebSocket.
 */

(function () {
  'use strict';

  /* ── Edit-item modal ──────────────────────────────────────────────────── */
  window.openEditModal = function (itemId, itemName, avail) {
    document.getElementById('edit-item-form').action = '/item/' + itemId + '/edit/';
    document.getElementById('edit-item-name').textContent  = itemName;
    document.getElementById('edit-avail').value            = avail;
    document.getElementById('edit-item-modal').style.display = 'flex';
  };

  window.closeEditModal = function () {
    document.getElementById('edit-item-modal').style.display = 'none';
  };

  document.addEventListener('DOMContentLoaded', function () {
    const modal = document.getElementById('edit-item-modal');
    if (modal) {
      modal.addEventListener('click', function (e) {
        if (e.target === this) closeEditModal();
      });
    }

    /* Sidebar badges via WS */
    const indicator = document.getElementById('rt-indicator');
    if (typeof InvSysRT !== 'undefined') {
      InvSysRT.connect('/ws/dashboard/', function (data) {
        if (typeof data.pending_count === 'number') {
          window.dispatchEvent(
            new CustomEvent('invsys:pending_count', { detail: data.pending_count })
          );
        }
        if (typeof data.graduation_warning_count === 'number') {
          window.dispatchEvent(
            new CustomEvent('invsys:grad_warning_count', { detail: data.graduation_warning_count })
          );
        }
      }, indicator);
    }
  });
})();