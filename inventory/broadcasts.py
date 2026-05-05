# inventory/broadcasts.py — replace the entire file

from asgiref.sync import async_to_sync
from channels.layers import get_channel_layer

from inventory.consumers import (
    _build_dashboard_payload,
    _build_borrow_management_payload,
    _build_borrow_requests_payload,
    _build_device_monitoring_payload,
)


def _send(group: str, payload: dict):
    layer = get_channel_layer()
    if layer is None:
        return
    async_to_sync(layer.group_send)(group, payload)


def _bust_caches():
    """Invalidate all server-side caches that nav or pages depend on."""
    from django.core.cache import cache
    from inventory.context_processors import BADGES_CACHE_KEY
    cache.delete_many([
        BADGES_CACHE_KEY,
        'dashboard_stats',
        'ajax_borrow_mgmt',
        'ajax_device_monitoring',
        'ajax_borrow_requests',
    ])


def broadcast_dashboard():
    _bust_caches()
    payload = _build_dashboard_payload()
    _send('dashboard', payload)


def broadcast_borrow_management():
    _bust_caches()
    payload = _build_borrow_management_payload()
    _send('borrow_management', payload)


def broadcast_borrow_requests():
    _bust_caches()
    payload = _build_borrow_requests_payload()
    _send('borrow_requests', payload)


def broadcast_device_monitoring():
    _bust_caches()
    payload = _build_device_monitoring_payload()
    _send('device_monitoring', payload)


def broadcast_all():
    _bust_caches()
    broadcast_dashboard()
    broadcast_borrow_management()
    broadcast_borrow_requests()
    broadcast_device_monitoring()