import io
import random
import pytz
from django.utils import timezone
from celery import shared_task
from .models import Item, DeviceMonitor, TransactionDevice, BorrowRequest, Transaction

PH_TZ = pytz.timezone('Asia/Manila')

def get_ph_time():
    return timezone.now().astimezone(PH_TZ)

def _parse_excel_date(raw):
    # ... (same as your original, copy it here)
    pass

@shared_task
def process_excel_import(rows_data, user_id):
    """
    rows_data: list of dicts (already parsed from the Excel file)
    user_id: ID of the staff user who initiated the import
    """
    from django.contrib.auth import get_user_model
    User = get_user_model()
    user = User.objects.get(pk=user_id)

    now_ph = get_ph_time()
    dummy_item, _ = Item.objects.get_or_create(name='Tablet (Import)', defaults={'quantity':0, 'available_quantity':0})

    serials = [d['serial_number'] for d in rows_data]
    existing = {dm.serial_number: dm for dm in DeviceMonitor.objects.filter(serial_number__in=serials)}

    to_create = []
    to_update = []
    returned_serials = []
    released_rows = []

    for d in rows_data:
        serial = d['serial_number']
        obj = existing.get(serial)

        accountable_person = d.get('accountable_person', '').strip()
        office_college = d.get('office_college', '').strip() or 'Unknown'
        borrower_type = 'employee' if d.get('borrower_type', '').lower() == 'employee' else 'student'
        is_returned = d.get('is_returned', False)
        is_released = d.get('is_released', False)
        date_returned = _parse_excel_date(d.get('date_returned_raw'))
        if is_returned and not date_returned:
            date_returned = now_ph
        if is_released:
            date_returned = None

        defaults = {
            'box_number': d.get('box_number', ''),
            'office_college': office_college,
            'accountable_person': accountable_person,
            'borrower_type': borrower_type,
            'accountable_officer': d.get('accountable_officer', ''),
            'assigned_mr': d.get('assigned_mr', ''),
            'device': d.get('device', '') or 'Tablet',
            'ptr': d.get('ptr', ''),
            'remarks': d.get('remarks', ''),
            'issue': d.get('issue', ''),
            'date_returned': date_returned,
            'is_released': is_released,
            'serviceable': False,
            'non_serviceable': False,
            'sealed': False,
            'missing': False,
            'incomplete': False,
        }

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

    if to_create:
        DeviceMonitor.objects.bulk_create(to_create)
    if to_update:
        update_fields = ['box_number','office_college','accountable_person','borrower_type','accountable_officer','assigned_mr','device','ptr','remarks','issue','date_returned','is_released','serviceable','non_serviceable','sealed','missing','incomplete']
        DeviceMonitor.objects.bulk_update(to_update, update_fields)

    if returned_serials:
        TransactionDevice.objects.filter(serial_number__in=returned_serials, returned=False).update(returned=True, returned_at=now_ph)

    if released_rows:
        # same bulk logic for BorrowRequest/Transaction/TransactionDevice as before
        released_serials_list = [d['serial_number'] for d in released_rows]
        TransactionDevice.objects.filter(serial_number__in=released_serials_list, returned=False).update(returned=True, returned_at=now_ph)

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
        for i, d in enumerate(released_rows):
            borrow_reqs.append(BorrowRequest(
                transaction_id=new_ids[i],
                borrower_name=d.get('accountable_person',''),
                borrower_type=d.get('borrower_type','student'),
                office_college=d.get('office_college','Unknown'),
                college=d.get('office_college','Unknown'),
                item=None, quantity=1, status='accepted',
                student_id='', year_level='', section='', academic_year='',
            ))
        borrow_reqs = BorrowRequest.objects.bulk_create(borrow_reqs)

        txs = []
        for i, d in enumerate(released_rows):
            txs.append(Transaction(
                borrow_request=borrow_reqs[i], item=dummy_item, borrower=user,
                office_college=d.get('office_college','Unknown'),
                quantity_borrowed=1, returned_qty=0, status='borrowed',
                borrowed_at=now_ph, serial_number=d['serial_number'],
            ))
        txs = Transaction.objects.bulk_create(txs)

        TransactionDevice.objects.bulk_create([
            TransactionDevice(transaction=txs[i], serial_number=d['serial_number'],
                              box_number=d.get('box_number',''), returned=False)
            for i, d in enumerate(released_rows)
        ])

    # After processing, broadcast update
    from inventory.broadcasts import broadcast_device_monitoring
    broadcast_device_monitoring()