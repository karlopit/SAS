"""
inventory/consumers.py
"""
import json
from channels.generic.websocket import AsyncWebsocketConsumer
from channels.db import database_sync_to_async


def _build_dashboard_payload():
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
        'offices':     offices,
        'serviceable': [monitors.filter(office_college=o, serviceable=True).count()     for o in offices],
        'nonService':  [monitors.filter(office_college=o, non_serviceable=True).count() for o in offices],
        'sealed':      [monitors.filter(office_college=o, sealed=True).count()          for o in offices],
        'missing':     [monitors.filter(office_college=o, missing=True).count()         for o in offices],
        'incomplete':  [monitors.filter(office_college=o, incomplete=True).count()      for o in offices],
    }

    return {
        'type':           'dashboard.update',
        'items_count':    items.count(),
        'active_borrows': active_borrows,
        'total_returns':  total_returns,
        'pending_count':  pending_count,
        'available_qty':  available_qty,
        'borrowed_qty':   borrowed_qty,
        'bar':            bar,
    }


# In consumers.py, update the _build_borrow_management_payload function:

def _build_borrow_management_payload():
    from .models import Transaction, BorrowRequest, Item
    from .views import format_ph_time  # Import the formatting function
    
    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).order_by('-borrowed_at')[:50]
    
    transactions_data = []
    for tx in transactions:
        # Get borrower name
        if tx.borrow_request:
            borrower_name = tx.borrow_request.borrower_name
            borrower_type = tx.borrow_request.borrower_type
            tx_id = tx.borrow_request.transaction_id
        else:
            borrower_name = tx.borrower.username
            borrower_type = ''
            tx_id = ''
        
        # Get accountable officer
        accountable_officer = tx.borrower.get_full_name() or tx.borrower.username
        
        transactions_data.append({
            'id': tx.id,
            'tx_id': tx_id,
            'borrower_name': borrower_name,
            'borrower_type': borrower_type,
            'accountable_officer': accountable_officer,
            'office_college': tx.office_college or '',
            'item_name': tx.item.name,
            'qty_borrowed': tx.quantity_borrowed,
            'returned_qty': tx.returned_qty,
            'borrowed_at': format_ph_time(tx.borrowed_at),  # ← Use formatted time
            'returned_at': format_ph_time(tx.returned_at) if tx.returned_at else '—',  # ← Use formatted time
            'fully_returned': tx.returned_qty >= tx.quantity_borrowed,
        })
    
    items_data = []
    for item in Item.objects.all():
        items_data.append({
            'id': item.id,
            'name': item.name,
            'serial': item.serial,
            'description': item.description,
            'quantity': item.quantity,
            'available_quantity': item.available_quantity,
        })
    
    return {
        'type': 'borrow_management.update',
        'transactions': transactions_data,
        'items': items_data,
        'pending_count': BorrowRequest.objects.filter(status='pending').count(),
    }


def _build_borrow_requests_payload():
    from inventory.models import BorrowRequest

    pending_qs = BorrowRequest.objects.filter(
        status='pending'
    ).select_related('item').order_by('-created_at')

    pending = []
    for r in pending_qs:
        pending.append({
            'id':             r.id,
            'transaction_id': r.transaction_id,
            'borrower_name':  r.borrower_name,
            'office_college': r.office_college,
            'item_name':      r.item.name if r.item else '—',
            'quantity':       r.quantity,
            'created_at':     r.created_at.strftime('%b %d, %Y — %H:%M'),
        })

    return {
        'type':    'borrow_requests.update',
        'pending': pending,
        'count':   len(pending),
    }


def _build_device_monitoring_payload():
    from inventory.models import DeviceMonitor, TransactionDevice
    from .views import format_ph_time

    rows_qs = DeviceMonitor.objects.all().order_by('id')
    rows = []
    for row in rows_qs:
        if row.date_returned:
            release_status = 'Returned'
            date_returned_str = format_ph_time(row.date_returned)
        else:
            active_td = TransactionDevice.objects.filter(
                serial_number=row.serial_number,
                returned=False
            ).select_related('transaction').first()
            if active_td and active_td.transaction:
                tx = active_td.transaction
                tx_borrower = tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username
                if tx_borrower == row.accountable_person and tx.office_college == row.office_college:
                    release_status = 'Released'
                else:
                    release_status = '—'
            else:
                release_status = '—'
            date_returned_str = '—'

        rows.append({
            'id':                 row.id,
            'box_number':         row.box_number,
            'office_college':     row.office_college,
            'accountable_person': row.accountable_person,
            'borrower_type':      row.borrower_type,
            'accountable_officer': row.accountable_officer,
            'device':             row.device,
            'serial_number':      row.serial_number,
            'serviceable':        row.serviceable,
            'non_serviceable':    row.non_serviceable,
            'sealed':             row.sealed,
            'missing':            row.missing,
            'incomplete':         row.incomplete,
            'remarks':            row.remarks,
            'issue':              row.issue,
            'release_status':     release_status,
            'date_returned':      date_returned_str,
        })

    return {
        'type': 'device_monitoring.update',
        'rows': rows,
    }


# ── Base Consumer ─────────────────────────────────────────────────────────────

class BaseConsumer(AsyncWebsocketConsumer):
    group_name = None

    async def connect(self):
        if not self.scope['user'].is_authenticated:
            await self.close()
            return
        await self.channel_layer.group_add(self.group_name, self.channel_name)
        await self.accept()
        payload = await database_sync_to_async(self.build_payload)()
        await self.send(text_data=json.dumps(payload))

    async def disconnect(self, code):
        await self.channel_layer.group_discard(self.group_name, self.channel_name)

    def build_payload(self):
        raise NotImplementedError

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