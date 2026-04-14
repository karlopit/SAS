"""
inventory/consumers.py

WebSocket consumers that push live data to connected clients.

Groups used:
  dashboard          — all authenticated users on the dashboard
  borrow_management  — staff on the Borrow Management page
  borrow_requests    — staff on the Borrow Requests page
  device_monitoring  — staff on Device Monitoring
"""
import json
from channels.generic.websocket import AsyncWebsocketConsumer
from channels.db import database_sync_to_async
from django.utils import timezone


# ── helpers ──────────────────────────────────────────────────────────────────

def _build_dashboard_payload():
    """Synchronous DB read — called via database_sync_to_async."""
    from django.db.models import Sum, F, ExpressionWrapper, IntegerField
    from inventory.models import Item, Transaction, BorrowRequest, DeviceMonitor

    items          = Item.objects.all()
    active_borrows = Transaction.objects.filter(status='borrowed').count()
    total_returns  = Transaction.objects.filter(status='returned').count()
    pending_count  = BorrowRequest.objects.filter(status='pending').count()
    available_qty  = sum(i.available_quantity for i in items)

    agg = Transaction.objects.annotate(
        still_out=ExpressionWrapper(
            F('quantity_borrowed') - F('returned_qty'),
            output_field=IntegerField()
        )
    ).aggregate(total=Sum('still_out'))
    borrowed_qty = max(0, agg['total'] or 0)

    monitors = DeviceMonitor.objects.all()
    offices  = sorted(set(monitors.values_list('office_college', flat=True)))

    bar = {
        'offices':      offices,
        'serviceable':  [monitors.filter(office_college=o, serviceable=True).count()     for o in offices],
        'nonService':   [monitors.filter(office_college=o, non_serviceable=True).count() for o in offices],
        'sealed':       [monitors.filter(office_college=o, sealed=True).count()          for o in offices],
        'missing':      [monitors.filter(office_college=o, missing=True).count()         for o in offices],
        'incomplete':   [monitors.filter(office_college=o, incomplete=True).count()      for o in offices],
    }

    return {
        'type':          'dashboard.update',
        'items_count':   items.count(),
        'active_borrows': active_borrows,
        'total_returns':  total_returns,
        'pending_count':  pending_count,
        'available_qty':  available_qty,
        'borrowed_qty':   borrowed_qty,
        'bar':            bar,
    }


def _build_borrow_management_payload():
    from inventory.models import Item, Transaction, BorrowRequest

    items = list(Item.objects.values(
        'id', 'name', 'serial', 'description', 'available_quantity', 'quantity'
    ))
    pending_count = BorrowRequest.objects.filter(status='pending').count()

    txs_qs = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).order_by('-borrowed_at')[:20]

    transactions = []
    for tx in txs_qs:
        transactions.append({
            'id':              tx.id,
            'tx_id':           f'#{tx.borrow_request.transaction_id}' if tx.borrow_request else '—',
            'borrower_name':   tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username,
            'office_college':  tx.office_college or '—',
            'item_name':       tx.item.name,
            'item_serial':     tx.item.serial or '—',
            'qty_borrowed':    tx.quantity_borrowed,
            'returned_qty':    tx.returned_qty,
            'borrowed_at':     tx.borrowed_at.strftime('%b %d, %Y'),
            'returned_at':     tx.returned_at.strftime('%b %d, %Y %H:%M') if tx.returned_at else '—',
            'fully_returned':  tx.returned_qty >= tx.quantity_borrowed,
            'status':          tx.status,
        })

    return {
        'type':          'borrow_management.update',
        'items':         items,
        'transactions':  transactions,
        'pending_count': pending_count,
    }


def _build_borrow_requests_payload():
    from inventory.models import BorrowRequest

    pending_qs = BorrowRequest.objects.filter(
        status='pending'
    ).select_related('item').order_by('-created_at')

    pending = []
    for r in pending_qs:
        pending.append({
            'id':            r.id,
            'transaction_id': r.transaction_id,
            'borrower_name': r.borrower_name,
            'office_college': r.office_college,
            'item_name':     r.item.name if r.item else '—',
            'quantity':      r.quantity,
            'created_at':    r.created_at.strftime('%b %d, %Y — %H:%M'),
        })

    return {
        'type':    'borrow_requests.update',
        'pending': pending,
        'count':   len(pending),
    }


def _build_device_monitoring_payload():
    from inventory.models import DeviceMonitor

    rows_qs = DeviceMonitor.objects.all().order_by('id')
    rows = []
    for r in rows_qs:
        rows.append({
            'id':                 r.id,
            'display_id':         r.display_id,
            'office_college':     r.office_college,
            'accountable_person': r.accountable_person,
            'device':             r.device,
            'serial_number':      r.serial_number,
            'serviceable':        r.serviceable,
            'non_serviceable':    r.non_serviceable,
            'sealed':             r.sealed,
            'missing':            r.missing,
            'incomplete':         r.incomplete,
        })

    return {
        'type': 'device_monitoring.update',
        'rows': rows,
    }


# ── Base Consumer ─────────────────────────────────────────────────────────────

class BaseConsumer(AsyncWebsocketConsumer):
    """
    Shared connect/disconnect logic.
    Sub-classes set `group_name` and override `build_payload()`.
    """
    group_name = None

    async def connect(self):
        if not self.scope['user'].is_authenticated:
            await self.close()
            return
        await self.channel_layer.group_add(self.group_name, self.channel_name)
        await self.accept()
        # Send initial snapshot immediately on connect
        payload = await database_sync_to_async(self.build_payload)()
        await self.send(text_data=json.dumps(payload))

    async def disconnect(self, code):
        await self.channel_layer.group_discard(self.group_name, self.channel_name)

    def build_payload(self):
        raise NotImplementedError

    # Generic group-broadcast handler — message type must match group name pattern
    async def _broadcast(self, event):
        await self.send(text_data=json.dumps(event))


# ── Concrete Consumers ────────────────────────────────────────────────────────

class DashboardConsumer(BaseConsumer):
    group_name = 'dashboard'

    def build_payload(self):
        return _build_dashboard_payload()

    async def dashboard_update(self, event):
        await self._broadcast(event)


class BorrowManagementConsumer(BaseConsumer):
    group_name = 'borrow_management'

    def build_payload(self):
        return _build_borrow_management_payload()

    async def borrow_management_update(self, event):
        await self._broadcast(event)


class BorrowRequestsConsumer(BaseConsumer):
    group_name = 'borrow_requests'

    def build_payload(self):
        return _build_borrow_requests_payload()

    async def borrow_requests_update(self, event):
        await self._broadcast(event)


class DeviceMonitoringConsumer(BaseConsumer):
    group_name = 'device_monitoring'

    def build_payload(self):
        return _build_device_monitoring_payload()

    async def device_monitoring_update(self, event):
        await self._broadcast(event)