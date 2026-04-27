import json
import pytz
from channels.generic.websocket import AsyncWebsocketConsumer
from channels.db import database_sync_to_async

PH_TZ = pytz.timezone('Asia/Manila')

def _fmt_ph(dt):
    """Format a datetime to Philippine time — matches format_ph_time() in views.py."""
    if not dt:
        return '—'
    import django.utils.timezone as tz
    if tz.is_naive(dt):
        dt = tz.make_aware(dt, tz.utc)
    return dt.astimezone(PH_TZ).strftime('%b %d, %Y %I:%M %p')


def _build_dashboard_payload():
    from django.db.models import Sum, F, ExpressionWrapper, IntegerField
    from inventory.models import Item, Transaction, BorrowRequest, DeviceMonitor

    items_count   = Item.objects.count()
    available_qty = Item.objects.aggregate(t=Sum('available_quantity'))['t'] or 0

    # Two fast index-scan counts — no broken boolean aggregate
    active_borrows = Transaction.objects.filter(status='borrowed').count()
    total_returns  = Transaction.objects.filter(status='returned').count()
    pending_count  = BorrowRequest.objects.filter(status='pending').count()

    out_agg = Transaction.objects.annotate(
        still_out=ExpressionWrapper(
            F('quantity_borrowed') - F('returned_qty'),
            output_field=IntegerField()
        )
    ).aggregate(total=Sum('still_out'))
    borrowed_qty = max(0, out_agg['total'] or 0)

    # Bar chart — one bulk query per status flag, grouped by office
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

    # ── Graduation warning count (same logic as context processor) ──
    graduating_keywords = ['4th', 'fourth', '5th', 'fifth']
    active_trans = Transaction.objects.filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).select_related('borrow_request')
    grad_count = 0
    for tx in active_trans:
        br = tx.borrow_request
        if br:
            yl = (br.year_level or br.year_section or '').strip().lower()
            if any(k in yl for k in graduating_keywords):
                grad_count += 1

    return {
        'type':           'dashboard.update',
        'items_count':    items_count,
        'active_borrows': active_borrows,
        'total_returns':  total_returns,
        'pending_count':  pending_count,
        'available_qty':  available_qty,
        'borrowed_qty':   borrowed_qty,
        'bar':            bar,
        'graduation_warning_count': grad_count,   # <── new field
    }


def _build_borrow_management_payload():
    from .models import Transaction, BorrowRequest, Item

    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).order_by('-borrowed_at')[:50]

    transactions_data = []
    for tx in transactions:
        if tx.borrow_request:
            borrower_name = tx.borrow_request.borrower_name
            borrower_type = tx.borrow_request.borrower_type
            tx_id         = tx.borrow_request.transaction_id
        else:
            borrower_name = tx.borrower.username
            borrower_type = ''
            tx_id         = ''

        accountable_officer = (tx.borrower.get_full_name() or '').strip() or tx.borrower.username

        transactions_data.append({
            'id':                  tx.id,
            'tx_id':               tx_id,
            'borrower_name':       borrower_name,
            'borrower_type':       borrower_type,
            'accountable_officer': accountable_officer,
            'office_college':      tx.office_college or '',
            'item_name':           tx.item.name,
            'qty_borrowed':        tx.quantity_borrowed,
            'returned_qty':        tx.returned_qty,
            'borrowed_at':         _fmt_ph(tx.borrowed_at),
            'returned_at':         _fmt_ph(tx.returned_at) if tx.returned_at else '—',
            'fully_returned':      tx.returned_qty >= tx.quantity_borrowed,
        })

    items_data = list(Item.objects.values(
        'id', 'name', 'serial', 'description', 'quantity', 'available_quantity'
    ))

    return {
        'type':          'borrow_management.update',
        'transactions':  transactions_data,
        'items':         items_data,
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
    """
    Build device monitoring payload.
    Uses a single bulk lookup instead of per-row DB queries.
    """
    from inventory.models import DeviceMonitor, Transaction, TransactionDevice

    active_tds = TransactionDevice.objects.filter(
        returned=False
    ).select_related('transaction', 'transaction__borrow_request').values(
        'serial_number',
        'transaction__returned_qty',
        'transaction__quantity_borrowed',
        'transaction__returned_at',
        'transaction__borrow_request__borrower_name',
        'transaction__office_college',
    )

    active_serial_map = {}
    for td in active_tds:
        sn = td['serial_number']
        if sn and sn not in active_serial_map:
            active_serial_map[sn] = td

    serial_to_tx = {}
    for tx in Transaction.objects.select_related('borrow_request').order_by('-borrowed_at'):
        if not tx.serial_number:
            continue
        for sn in [s.strip() for s in tx.serial_number.split(',') if s.strip()]:
            if sn not in serial_to_tx:
                serial_to_tx[sn] = tx

    rows_qs = DeviceMonitor.objects.all().order_by('id')
    rows = []
    for r in rows_qs:
        sn = (r.serial_number or '').strip()

        if r.date_returned:
            release_status    = 'Returned'
            date_returned_str = _fmt_ph(r.date_returned)
        elif sn and sn in active_serial_map:
            td_data = active_serial_map[sn]
            tx_borrower = td_data['transaction__borrow_request__borrower_name'] or ''
            tx_office   = td_data['transaction__office_college'] or ''

            if tx_borrower == r.accountable_person and tx_office == r.office_college:
                release_status = 'Released'
            else:
                release_status = '—'
            date_returned_str = '—'
        elif sn and sn in serial_to_tx:
            tx = serial_to_tx[sn]
            release_status    = 'Returned' if tx.returned_qty >= tx.quantity_borrowed else 'Released'
            date_returned_str = _fmt_ph(tx.returned_at) if tx.returned_at else '—'
        else:
            release_status    = '—'
            date_returned_str = '—'

        rows.append({
            'id':                  r.id,
            'box_number':          r.box_number,
            'office_college':      r.office_college,
            'accountable_person':  r.accountable_person,
            'borrower_type':       r.borrower_type,
            'accountable_officer': r.accountable_officer,
            'device':              r.device,
            'serial_number':       r.serial_number,
            'serviceable':         r.serviceable,
            'non_serviceable':     r.non_serviceable,
            'sealed':              r.sealed,
            'missing':             r.missing,
            'incomplete':          r.incomplete,
            'remarks':             r.remarks,
            'issue':               r.issue,
            'release_status':      release_status,
            'date_returned':       date_returned_str,
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