/**
 * realtime.js — WebSocket client with AJAX-poll fallback
 *
 * Usage (in a page template):
 *   InvSysRT.connect('/ws/dashboard/', onMessage, indicatorElement);
 *
 * The module:
 *   1. Opens a WebSocket.
 *   2. On every message it calls onMessage(data).
 *   3. If the WS fails / is unavailable it falls back to polling the
 *      matching /ajax/… endpoint every POLL_INTERVAL ms.
 *   4. Dispatches custom events for badge updates ONLY when the value
 *      is a number (including 0). This prevents blank badges.
 */

(function (global) {
  'use strict';

  const POLL_INTERVAL = 5000;   // ms between AJAX polls
  const WS_RETRY_MS   = 3000;   // ms before attempting WS reconnect

  /** Map WS path → AJAX fallback URL */
  const WS_TO_AJAX = {
    '/ws/dashboard/':          '/ajax/dashboard/',
    '/ws/borrow-management/':  '/ajax/borrow-management/',
    '/ws/borrow-requests/':    '/ajax/borrow-requests/',
    '/ws/device-monitoring/':  '/ajax/device-monitoring/',
  };

  /* ── Status indicator helpers ─────────────────────────────────────────── */
  function _setIndicator(el, state) {
    if (!el) return;
    el.className = `rt-indicator rt-${state}`;
    const labels = { connected: 'Live', polling: 'Polling', disconnected: 'Offline' };
    el.textContent = labels[state] ?? state;
  }

  /* ── Custom event dispatcher for live badge updates ───────────────────── */
  function _dispatchEvents(data) {
    // Only fire if the value exists and is a number (including 0)
    if (typeof data.pending_count === 'number') {
      window.dispatchEvent(new CustomEvent('invsys:pending_count', {
        detail: data.pending_count
      }));
    }
    if (typeof data.graduation_warning_count === 'number') {
      window.dispatchEvent(new CustomEvent('invsys:grad_warning_count', {
        detail: data.graduation_warning_count
      }));
    }
  }

  /* ── WebSocket transport ──────────────────────────────────────────────── */
  function _wsUrl(path) {
    const proto = location.protocol === 'https:' ? 'wss' : 'ws';
    return `${proto}://${location.host}${path}`;
  }

  /* ── Main connect function ────────────────────────────────────────────── */
  /**
   * @param {string}   wsPath    e.g. '/ws/dashboard/'
   * @param {Function} onMessage called with parsed JSON data on every update
   * @param {Element}  [indicator] optional DOM element showing connection state
   * @returns {{ close: Function }} handle
   */
  function connect(wsPath, onMessage, indicator) {
    let ws          = null;
    let pollTimer   = null;
    let retryTimer  = null;
    let closed      = false;

    const ajaxUrl = WS_TO_AJAX[wsPath];

    /* ── AJAX poll ── */
    function startPolling() {
      if (pollTimer) return;
      _setIndicator(indicator, 'polling');
      _poll();
      pollTimer = setInterval(_poll, POLL_INTERVAL);
    }

    function stopPolling() {
      clearInterval(pollTimer);
      pollTimer = null;
    }

    function _poll() {
      if (!ajaxUrl) return;
      fetch(ajaxUrl, { credentials: 'same-origin' })
        .then(r => r.ok ? r.json() : Promise.reject(r.status))
        .then(data => {
          onMessage(data);
          _dispatchEvents(data);
        })
        .catch(err => console.warn('[InvSysRT] AJAX poll error:', err));
    }

    /* ── WebSocket ── */
    function openWS() {
      if (closed) return;
      try {
        ws = new WebSocket(_wsUrl(wsPath));
      } catch (e) {
        startPolling();
        return;
      }

      ws.onopen = () => {
        stopPolling();
        _setIndicator(indicator, 'connected');
        clearTimeout(retryTimer);
      };

      ws.onmessage = (event) => {
        try {
          const data = JSON.parse(event.data);
          onMessage(data);
          _dispatchEvents(data);
        } catch (e) {
          console.warn('[InvSysRT] Parse error:', e);
        }
      };

      ws.onerror = () => {
        /* intentionally silent — onclose will handle fallback */
      };

      ws.onclose = (ev) => {
        if (closed) return;
        _setIndicator(indicator, 'disconnected');
        startPolling();          // start polling while we wait to reconnect
        retryTimer = setTimeout(() => {
          stopPolling();
          openWS();
        }, WS_RETRY_MS);
      };
    }

    openWS();

    return {
      close() {
        closed = true;
        stopPolling();
        clearTimeout(retryTimer);
        if (ws) ws.close();
      }
    };
  }

  global.InvSysRT = { connect };
})(window);