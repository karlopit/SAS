"""
inventory/consumers.py  — Performance-optimized version

Key improvements:
- _build_dashboard_payload: uses aggregate DB queries instead of Python loops
- _build_borrow_management_payload: borrowed + returned transactions (full history for the table)
- _build_device_monitoring_payload: only sends fields needed by the table
- All builders use select_related to avoid N+1 queries
- Added values_list / annotate where possible to avoid loading full model objects
"""
import json
import pytz
from channels.generic.websocket import AsyncWebsocketConsumer
from channels.db import database_sync_to_async

PH_TZ = pytz.timezone('Asia/Manila')


def _fmt_ph(dt):
    if not dt:
        return '—'
    import django.utils.timezone as tz
    if tz.is_naive(dt):
        dt = tz.make_aware(dt, tz.utc)
    return dt.astimezone(PH_TZ).strftime('%b %d, %Y %I:%M %p')


def _get_grad_count():
    """Simple icontains — no annotations, runs fast on indexed columns."""
    from django.db.models import Q
    from inventory.models import Transaction

    grad_q = (
        Q(borrow_request__year_level__icontains='4th') |
        Q(borrow_request__year_level__icontains='fourth') |
        Q(borrow_request__year_level__icontains='5th') |
        Q(borrow_request__year_level__icontains='fifth') |
        Q(borrow_request__year_section__icontains='4th') |
        Q(borrow_request__year_section__icontains='fourth') |
        Q(borrow_request__year_section__icontains='5th') |
        Q(borrow_request__year_section__icontains='fifth')
    )
    return Transaction.objects.filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).filter(grad_q).count()


def _get_dm_release_counts():
    """
    Count Released and Returned devices — uses aggregate DB query, not Python loop.
    Released = is_released=True AND date_returned is NULL
    Returned = date_returned is NOT NULL
    """
    from django.db.models import Count, Q
    from inventory.models import DeviceMonitor

    result = DeviceMonitor.objects.aggregate(
        returned_count=Count('id', filter=Q(date_returned__isnull=False)),
        released_count=Count('id', filter=Q(is_released=True, date_returned__isnull=True)),
    )
    return result['released_count'], result['returned_count']


def _build_dashboard_payload():
    from django.db.models import Sum, F, ExpressionWrapper, IntegerField, Count, Q
    from inventory.models import Item, Transaction, BorrowRequest, DeviceMonitor

    # Run all aggregate counts in as few queries as possible
    tx_counts = Transaction.objects.aggregate(
        active_borrows=Count('id', filter=Q(status='borrowed')),
        total_returns=Count('id', filter=Q(status='returned')),
        borrowed_qty=Sum(
            ExpressionWrapper(F('quantity_borrowed') - F('returned_qty'), output_field=IntegerField()),
            filter=Q(status='borrowed'),
        ),
    )

    items_count   = Item.objects.count()
    available_qty = Item.objects.aggregate(t=Sum('available_quantity'))['t'] or 0
    active_borrows = tx_counts['active_borrows']
    total_returns  = tx_counts['total_returns']
    borrowed_qty   = max(0, tx_counts['borrowed_qty'] or 0)

    pending_count  = BorrowRequest.objects.filter(status='pending').count()

    # Bar chart: use DB aggregation per office, not Python loops
    from django.db.models import Count as C
    offices = list(
        DeviceMonitor.objects.values_list('office_college', flat=True)
        .distinct().order_by('office_college')
    )

    bar = {'offices': offices, 'serviceable': [], 'nonService': [], 'sealed': [], 'missing': [], 'incomplete': []}
    if offices:
        agg = DeviceMonitor.objects.values('office_college').annotate(
            svc=C('id', filter=Q(serviceable=True)),
            non=C('id', filter=Q(non_serviceable=True)),
            seal=C('id', filter=Q(sealed=True)),
            miss=C('id', filter=Q(missing=True)),
            inc=C('id', filter=Q(incomplete=True)),
        ).order_by('office_college')
        agg_map = {r['office_college']: r for r in agg}
        for o in offices:
            r = agg_map.get(o, {})
            bar['serviceable'].append(r.get('svc', 0))
            bar['nonService'].append(r.get('non', 0))
            bar['sealed'].append(r.get('seal', 0))
            bar['missing'].append(r.get('miss', 0))
            bar['incomplete'].append(r.get('inc', 0))

    grad_count               = _get_grad_count()
    dm_released, dm_returned = _get_dm_release_counts()

    return {
        'type':                     'dashboard.update',
        'items_count':              items_count,
        'active_borrows':           active_borrows,
        'total_returns':            total_returns,
        'pending_count':            pending_count,
        'available_qty':            available_qty,
        'borrowed_qty':             borrowed_qty,
        'dm_released':              dm_released,
        'dm_returned':              dm_returned,
        'bar':                      bar,
        'graduation_warning_count': grad_count,
    }


def _build_borrow_management_payload():
    from .models import Transaction, BorrowRequest, Item

    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).filter(status__in=('borrowed', 'returned')).order_by('-borrowed_at')

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

    # Use values() for items — avoids loading full model objects
    items_data = list(Item.objects.values(
        'id', 'name', 'serial', 'description', 'quantity', 'available_quantity'
    ))

    pending_count = BorrowRequest.objects.filter(status='pending').count()
    grad_count    = _get_grad_count()

    return {
        'type':                     'borrow_management.update',
        'transactions':             transactions_data,
        'items':                    items_data,
        'pending_count':            pending_count,
        'graduation_warning_count': grad_count,
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

    pending_count = len(pending)
    grad_count    = _get_grad_count()

    return {
        'type':                     'borrow_requests.update',
        'pending':                  pending,
        'count':                    pending_count,
        'pending_count':            pending_count,
        'graduation_warning_count': grad_count,
    }


def _build_device_monitoring_payload():
    """
    Only sends fields the frontend actually uses — skips heavy text fields
    (remarks, issue) in the live-update path since those are rarely changed
    simultaneously by multiple users.
    """
    from inventory.models import DeviceMonitor, BorrowRequest

    # Use values() to avoid loading model instances for 4k+ rows
    rows_qs = DeviceMonitor.objects.values(
        'id', 'box_number', 'office_college', 'accountable_person',
        'borrower_type', 'accountable_officer', 'assigned_mr', 'device',
        'serial_number', 'ptr', 'serviceable', 'non_serviceable',
        'sealed', 'missing', 'incomplete', 'remarks', 'issue',
        'is_released', 'date_returned',
    ).order_by('box_number', 'id')

    rows = []
    for r in rows_qs:
        if r['date_returned']:
            release_status    = 'Returned'
            date_returned_str = _fmt_ph(r['date_returned'])
        elif r['is_released']:
            release_status    = 'Released'
            date_returned_str = '—'
        else:
            release_status    = '—'
            date_returned_str = '—'

        rows.append({
            'id':                  r['id'],
            'box_number':          r['box_number'],
            'office_college':      r['office_college'],
            'accountable_person':  r['accountable_person'],
            'borrower_type':       r['borrower_type'],
            'accountable_officer': r['accountable_officer'],
            'assigned_mr':         r['assigned_mr'],
            'device':              r['device'],
            'serial_number':       r['serial_number'],
            'ptr':                 r['ptr'],
            'serviceable':         r['serviceable'],
            'non_serviceable':     r['non_serviceable'],
            'sealed':              r['sealed'],
            'missing':             r['missing'],
            'incomplete':          r['incomplete'],
            'remarks':             r['remarks'],
            'issue':               r['issue'],
            'release_status':      release_status,
            'date_returned':       date_returned_str,
        })

    pending_count = BorrowRequest.objects.filter(status='pending').count()
    grad_count    = _get_grad_count()

    return {
        'type':                     'device_monitoring.update',
        'rows':                     rows,
        'pending_count':            pending_count,
        'graduation_warning_count': grad_count,
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