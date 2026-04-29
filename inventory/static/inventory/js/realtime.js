/**
 * realtime.js — WebSocket client with AJAX-poll fallback
 *
 * IMPORTANT: This module intentionally does NOT dispatch badge update events.
 * Each page's own handleMessage function dispatches invsys:pending_count and
 * invsys:grad_warning_count after receiving the full payload — this avoids
 * race conditions where partial/missing fields could wipe the badges.
 */
(function (global) {
  'use strict';

  const POLL_INTERVAL = 5000;
  const WS_RETRY_MS   = 3000;

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

  /* ── WebSocket URL builder ────────────────────────────────────────────── */
  function _wsUrl(path) {
    const proto = location.protocol === 'https:' ? 'wss' : 'ws';
    return `${proto}://${location.host}${path}`;
  }

  /* ── Main connect function ────────────────────────────────────────────── */
  function connect(wsPath, onMessage, indicator) {
    let ws         = null;
    let pollTimer  = null;
    let retryTimer = null;
    let closed     = false;

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
          // Let the page's own handler do everything, including badge dispatch
          onMessage(data);
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
          // Let the page's own handler do everything, including badge dispatch
          onMessage(data);
        } catch (e) {
          console.warn('[InvSysRT] Parse error:', e);
        }
      };

      ws.onerror = () => { /* silent — onclose handles fallback */ };

      ws.onclose = () => {
        if (closed) return;
        _setIndicator(indicator, 'disconnected');
        startPolling();
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