"""
inventory/tasks.py

Celery task for async Excel import of DeviceMonitor records.
Handles create/update per serial number, marks returned/released devices,
and broadcasts live updates when done.
"""
import random
import pytz
from datetime import datetime, date as _date

from celery import shared_task
from django.utils import timezone


PH_TZ = pytz.timezone('Asia/Manila')


def _get_ph_time():
    return timezone.now().astimezone(PH_TZ)


def _parse_excel_date(raw):
    """
    Convert an openpyxl cell value to a PH-timezone-aware datetime.
    Returns None for blank / unparseable values.
    """
    if raw is None or str(raw).strip() in ('', '—', '-', 'N/A', 'None'):
        return None
    if isinstance(raw, datetime):
        return raw if raw.tzinfo else PH_TZ.localize(raw)
    if isinstance(raw, _date) and not isinstance(raw, datetime):
        return PH_TZ.localize(datetime(raw.year, raw.month, raw.day))
    text = str(raw).strip()
    for fmt in (
        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
        '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %H:%M', '%m/%d/%Y',
        '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y',
        '%b %d, %Y %I:%M %p', '%b %d, %Y',
        '%B %d, %Y %I:%M %p', '%B %d, %Y',
    ):
        try:
            return PH_TZ.localize(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return None


@shared_task(bind=True)
def process_excel_import(self, rows_data, user_id):
    """
    rows_data : list of dicts already parsed from the Excel file by the view.
    user_id   : PK of the staff user who triggered the import.

    Each dict has keys:
        serial_number, box_number, office_college, accountable_person,
        borrower_type, accountable_officer, assigned_mr, device, ptr,
        remarks, issue,
        is_returned (bool), is_released (bool),
        date_returned_raw (raw cell value – passed through _parse_excel_date here)
    """
    from django.contrib.auth import get_user_model
    from inventory.models import (
        Item, DeviceMonitor, TransactionDevice,
        BorrowRequest, Transaction,
    )

    User = get_user_model()
    try:
        user = User.objects.get(pk=user_id)
    except User.DoesNotExist:
        return {'ok': False, 'error': f'User {user_id} not found'}

    now_ph = _get_ph_time()

    # Ensure a dummy item exists for "released via import" transactions
    dummy_item, _ = Item.objects.get_or_create(
        name='Tablet (Import)',
        defaults={'quantity': 0, 'available_quantity': 0},
    )

    # ── Batch-fetch existing DeviceMonitor rows ──────────────────────────────
    serials = [d['serial_number'] for d in rows_data if d.get('serial_number')]
    existing_map = {
        dm.serial_number: dm
        for dm in DeviceMonitor.objects.filter(serial_number__in=serials)
    }

    to_create = []
    to_update = []
    returned_serials = []
    released_rows = []
    errors = []

    for d in rows_data:
        serial = (d.get('serial_number') or '').strip()
        if not serial:
            continue

        is_returned = bool(d.get('is_returned'))
        is_released = bool(d.get('is_released'))
        date_returned = _parse_excel_date(d.get('date_returned_raw'))

        # If flagged returned but no date, use now
        if is_returned and not date_returned:
            date_returned = now_ph
        # If flagged released, clear any date_returned
        if is_released:
            date_returned = None

        defaults = {
            'box_number':          (d.get('box_number') or '').strip(),
            'office_college':      (d.get('office_college') or 'Unknown').strip(),
            'accountable_person':  (d.get('accountable_person') or '').strip(),
            'borrower_type':       (d.get('borrower_type') or '').strip().lower(),
            'accountable_officer': (d.get('accountable_officer') or '').strip(),
            'assigned_mr':         (d.get('assigned_mr') or '').strip(),
            'device':              (d.get('device') or 'Tablet').strip(),
            'ptr':                 (d.get('ptr') or '').strip(),
            'remarks':             (d.get('remarks') or '').strip(),
            'issue':               (d.get('issue') or '').strip(),
            'date_returned':       date_returned,
            'is_released':         is_released and not is_returned,
            # Checkboxes: always cleared on import (per spec)
            'serviceable':     False,
            'non_serviceable': False,
            'sealed':          False,
            'missing':         False,
            'incomplete':      False,
        }

        obj = existing_map.get(serial)
        if obj is None:
            to_create.append(DeviceMonitor(serial_number=serial, **defaults))
        else:
            for attr, val in defaults.items():
                setattr(obj, attr, val)
            to_update.append(obj)

        if is_returned:
            returned_serials.append(serial)
        elif is_released:
            released_rows.append(d)

    # ── Bulk DB writes ───────────────────────────────────────────────────────
    try:
        if to_create:
            DeviceMonitor.objects.bulk_create(to_create, ignore_conflicts=False)
        if to_update:
            update_fields = [
                'box_number', 'office_college', 'accountable_person', 'borrower_type',
                'accountable_officer', 'assigned_mr', 'device', 'ptr',
                'remarks', 'issue', 'date_returned', 'is_released',
                'serviceable', 'non_serviceable', 'sealed', 'missing', 'incomplete',
            ]
            DeviceMonitor.objects.bulk_update(to_update, update_fields)
    except Exception as exc:
        errors.append(f'Bulk write error: {exc}')

    # ── Mark returned devices in TransactionDevice ───────────────────────────
    if returned_serials:
        TransactionDevice.objects.filter(
            serial_number__in=returned_serials,
            returned=False,
        ).update(returned=True, returned_at=now_ph)

    # ── Create BorrowRequest + Transaction rows for Released devices ─────────
    if released_rows:
        released_serials_list = [d['serial_number'] for d in released_rows]

        # Mark any existing TransactionDevice rows as returned (they were re-released)
        TransactionDevice.objects.filter(
            serial_number__in=released_serials_list,
            returned=False,
        ).update(returned=True, returned_at=now_ph)

        # Generate unique 5-digit transaction IDs
        existing_ids = set(BorrowRequest.objects.values_list('transaction_id', flat=True))
        new_ids = []
        for _ in released_rows:
            while True:
                tx_id = str(random.randint(10000, 99999))
                if tx_id not in existing_ids:
                    existing_ids.add(tx_id)
                    new_ids.append(tx_id)
                    break

        borrow_reqs = BorrowRequest.objects.bulk_create([
            BorrowRequest(
                transaction_id=new_ids[i],
                borrower_name=d.get('accountable_person', '').strip(),
                borrower_type=(d.get('borrower_type') or 'student').strip().lower(),
                office_college=(d.get('office_college') or 'Unknown').strip(),
                college=(d.get('office_college') or 'Unknown').strip(),
                item=None,
                quantity=1,
                status='accepted',
                student_id='',
                year_level='',
                section='',
                academic_year='',
            )
            for i, d in enumerate(released_rows)
        ])

        txs = Transaction.objects.bulk_create([
            Transaction(
                borrow_request=borrow_reqs[i],
                item=dummy_item,
                borrower=user,
                office_college=(d.get('office_college') or 'Unknown').strip(),
                quantity_borrowed=1,
                returned_qty=0,
                status='borrowed',
                borrowed_at=now_ph,
                serial_number=d['serial_number'],
            )
            for i, d in enumerate(released_rows)
        ])

        TransactionDevice.objects.bulk_create([
            TransactionDevice(
                transaction=txs[i],
                serial_number=d['serial_number'],
                box_number=(d.get('box_number') or '').strip(),
                returned=False,
            )
            for i, d in enumerate(released_rows)
        ])

    # ── Live broadcast ───────────────────────────────────────────────────────
    try:
        from inventory.broadcasts import broadcast_device_monitoring, broadcast_dashboard
        broadcast_device_monitoring()
        broadcast_dashboard()
    except Exception:
        pass  # Don't fail the task just because broadcast errored

    return {
        'ok': True,
        'created': len(to_create),
        'updated': len(to_update),
        'errors':  errors,
    }