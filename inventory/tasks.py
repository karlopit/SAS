"""
inventory/tasks.py

Celery task for async Excel import of DeviceMonitor records.
Handles create/update per serial number, marks returned/released devices,
and broadcasts live updates when done.

Bulk operations are chunked at CHUNK_SIZE rows to prevent Neon Postgres
from timing out on large imports (4k+ rows).

MODIFIED: accountable_officer now defaults to the importing user's full name
         when the Excel column is empty or missing.
"""
import random
import pytz
from datetime import datetime, date as _date

from celery import shared_task
from django.utils import timezone


PH_TZ      = pytz.timezone('Asia/Manila')
CHUNK_SIZE = 500   # rows per bulk_create / bulk_update call


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


def _chunked_bulk_create(model, objects, chunk_size=CHUNK_SIZE, ignore_conflicts=False):
    """bulk_create in chunks to avoid Neon query size / timeout limits."""
    created = 0
    for i in range(0, len(objects), chunk_size):
        chunk = objects[i:i + chunk_size]
        model.objects.bulk_create(chunk, ignore_conflicts=ignore_conflicts)
        created += len(chunk)
    return created


def _chunked_bulk_update(model, objects, fields, chunk_size=CHUNK_SIZE):
    """bulk_update in chunks to avoid Neon query size / timeout limits."""
    updated = 0
    for i in range(0, len(objects), chunk_size):
        chunk = objects[i:i + chunk_size]
        model.objects.bulk_update(chunk, fields)
        updated += len(chunk)
    return updated


@shared_task(bind=True, soft_time_limit=300, time_limit=360)
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

    # ── Wake up Neon free tier before heavy queries ──────────────────────────
    try:
        from django.db import connection
        connection.ensure_connection()
    except Exception:
        pass

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

    to_create        = []
    to_update        = []
    returned_serials = []
    released_rows    = []
    errors           = []

    for d in rows_data:
        serial = (d.get('serial_number') or '').strip()
        if not serial:
            continue

        is_returned   = bool(d.get('is_returned'))
        is_released   = bool(d.get('is_released'))
        date_returned = _parse_excel_date(d.get('date_returned_raw'))

        # If flagged returned but no date, use now
        if is_returned and not date_returned:
            date_returned = now_ph
        # If flagged released, clear any date_returned
        if is_released:
            date_returned = None

        # ── Default accountable_officer to importing user if Excel is blank ──
        accountable_officer = (d.get('accountable_officer') or '').strip()
        if not accountable_officer:
            accountable_officer = user.get_full_name() or user.username

        defaults = {
            'box_number':          (d.get('box_number')          or '').strip(),
            'office_college':      (d.get('office_college')      or 'Unknown').strip(),
            'accountable_person':  (d.get('accountable_person')  or '').strip(),
            'borrower_type':       (d.get('borrower_type')       or '').strip().lower(),
            'accountable_officer': accountable_officer,   # <-- now defaults to the importing user
            'assigned_mr':         (d.get('assigned_mr')         or '').strip(),
            'device':              (d.get('device')              or 'Tablet').strip(),
            'ptr':                 (d.get('ptr')                 or '').strip(),
            'remarks':             (d.get('remarks')             or '').strip(),
            'issue':               (d.get('issue')               or '').strip(),
            'date_returned':       date_returned,
            'is_released':         is_released and not is_returned,
            # Checkboxes always cleared on import
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

    # ── Chunked bulk DB writes ───────────────────────────────────────────────
    update_fields = [
        'box_number', 'office_college', 'accountable_person', 'borrower_type',
        'accountable_officer', 'assigned_mr', 'device', 'ptr',
        'remarks', 'issue', 'date_returned', 'is_released',
        'serviceable', 'non_serviceable', 'sealed', 'missing', 'incomplete',
    ]

    try:
        if to_create:
            _chunked_bulk_create(DeviceMonitor, to_create)
        if to_update:
            _chunked_bulk_update(DeviceMonitor, to_update, update_fields)
    except Exception as exc:
        errors.append(f'Bulk write error: {exc}')

    # ── Mark returned devices in TransactionDevice ───────────────────────────
    if returned_serials:
        for i in range(0, len(returned_serials), CHUNK_SIZE):
            chunk = returned_serials[i:i + CHUNK_SIZE]
            TransactionDevice.objects.filter(
                serial_number__in=chunk,
                returned=False,
            ).update(returned=True, returned_at=now_ph)

    # ── Create BorrowRequest + Transaction rows for Released devices ─────────
    if released_rows:
        released_serials_list = [d['serial_number'] for d in released_rows]

        for i in range(0, len(released_serials_list), CHUNK_SIZE):
            chunk = released_serials_list[i:i + CHUNK_SIZE]
            TransactionDevice.objects.filter(
                serial_number__in=chunk,
                returned=False,
            ).update(returned=True, returned_at=now_ph)

        existing_ids = set(BorrowRequest.objects.values_list('transaction_id', flat=True))
        new_ids = []
        for _ in released_rows:
            while True:
                tx_id = str(random.randint(10000, 99999))
                if tx_id not in existing_ids:
                    existing_ids.add(tx_id)
                    new_ids.append(tx_id)
                    break

        borrow_reqs = []
        for i in range(0, len(released_rows), CHUNK_SIZE):
            chunk_rows = released_rows[i:i + CHUNK_SIZE]
            chunk_ids  = new_ids[i:i + CHUNK_SIZE]
            created = BorrowRequest.objects.bulk_create([
                BorrowRequest(
                    transaction_id=chunk_ids[j],
                    borrower_name=(chunk_rows[j].get('accountable_person') or '').strip(),
                    borrower_type=(chunk_rows[j].get('borrower_type') or 'student').strip().lower(),
                    office_college=(chunk_rows[j].get('office_college') or 'Unknown').strip(),
                    college=(chunk_rows[j].get('office_college') or 'Unknown').strip(),
                    item=None,
                    quantity=1,
                    status='accepted',
                    student_id='',
                    year_level='',
                    section='',
                    academic_year='',
                )
                for j in range(len(chunk_rows))
            ])
            borrow_reqs.extend(created)

        txs = []
        for i in range(0, len(released_rows), CHUNK_SIZE):
            chunk_rows = released_rows[i:i + CHUNK_SIZE]
            chunk_brs  = borrow_reqs[i:i + CHUNK_SIZE]
            created = Transaction.objects.bulk_create([
                Transaction(
                    borrow_request=chunk_brs[j],
                    item=dummy_item,
                    borrower=user,
                    office_college=(chunk_rows[j].get('office_college') or 'Unknown').strip(),
                    quantity_borrowed=1,
                    returned_qty=0,
                    status='borrowed',
                    borrowed_at=now_ph,
                    serial_number=chunk_rows[j]['serial_number'],
                )
                for j in range(len(chunk_rows))
            ])
            txs.extend(created)

        for i in range(0, len(released_rows), CHUNK_SIZE):
            chunk_rows = released_rows[i:i + CHUNK_SIZE]
            chunk_txs  = txs[i:i + CHUNK_SIZE]
            TransactionDevice.objects.bulk_create([
                TransactionDevice(
                    transaction=chunk_txs[j],
                    serial_number=chunk_rows[j]['serial_number'],
                    box_number=(chunk_rows[j].get('box_number') or '').strip(),
                    returned=False,
                )
                for j in range(len(chunk_rows))
            ])

    # ── Live broadcast ───────────────────────────────────────────────────────
    try:
        from inventory.broadcasts import broadcast_device_monitoring, broadcast_dashboard
        broadcast_device_monitoring()
        broadcast_dashboard()
    except Exception:
        pass

    return {
        'ok':      True,
        'created': len(to_create),
        'updated': len(to_update),
        'errors':  errors,
    }