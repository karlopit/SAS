"""
inventory/broadcasts.py

Call these after any mutation to push live updates to all connected WebSocket clients.
Uses async_to_sync so they can be called from synchronous Django views.

Usage inside a view:
    from inventory.broadcasts import broadcast_dashboard, broadcast_borrow_management
    broadcast_dashboard()
    broadcast_borrow_management()
"""
from asgiref.sync import async_to_sync
from channels.layers import get_channel_layer

from inventory.consumers import (
    _build_dashboard_payload,
    _build_borrow_management_payload,
    _build_borrow_requests_payload,
    _build_device_monitoring_payload,
)


def _send(group: str, payload: dict):
    """Push *payload* to every consumer in *group*."""
    layer = get_channel_layer()
    if layer is None:
        return
    async_to_sync(layer.group_send)(group, payload)


# ── Public API ────────────────────────────────────────────────────────────────

def broadcast_dashboard():
    payload = _build_dashboard_payload()
    _send('dashboard', payload)


def broadcast_borrow_management():
    payload = _build_borrow_management_payload()
    _send('borrow_management', payload)


def broadcast_borrow_requests():
    payload = _build_borrow_requests_payload()
    _send('borrow_requests', payload)


def broadcast_device_monitoring():
    payload = _build_device_monitoring_payload()
    _send('device_monitoring', payload)


def broadcast_all():
    """Convenience: push updates to every channel group at once."""
    broadcast_dashboard()
    broadcast_borrow_management()
    broadcast_borrow_requests()
    broadcast_device_monitoring()