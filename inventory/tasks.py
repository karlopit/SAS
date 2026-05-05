"""
inventory/tasks.py

Celery tasks for:
1. Async Excel import of DeviceMonitor records
2. Async Excel export of Borrow Management & Device Monitoring

All shared helpers come from inventory.utils to avoid circular imports.
"""
import io
import random
from datetime import datetime, date as _date
from celery import shared_task
from django.utils import timezone
from django.core.cache import cache
from inventory.utils import get_ph_time, format_ph_time, parse_excel_date, PH_TZ


CHUNK_SIZE = 500   # rows per bulk_create / bulk_update call


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


# ═══════════════════════════════════════════════════════════════════════════
#  IMPORT TASK
# ═══════════════════════════════════════════════════════════════════════════

@shared_task(bind=True, soft_time_limit=300, time_limit=360)
def process_excel_import(self, rows_data, user_id):
    """
    rows_data : list of dicts already parsed from the Excel file by the view.
    user_id   : PK of the staff user who triggered the import.
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

    # Wake up Neon free tier before heavy queries
    try:
        from django.db import connection
        connection.ensure_connection()
    except Exception:
        pass

    now_ph = get_ph_time()

    dummy_item, _ = Item.objects.get_or_create(
        name='Tablet (Import)',
        defaults={'quantity': 0, 'available_quantity': 0},
    )

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
        date_returned = parse_excel_date(d.get('date_returned_raw'))

        if is_returned and not date_returned:
            date_returned = now_ph
        if is_released:
            date_returned = None

        # Default accountable_officer to importing user if Excel is blank
        accountable_officer = (d.get('accountable_officer') or '').strip()
        if not accountable_officer:
            accountable_officer = user.get_full_name() or user.username

        # Clean box number: remove trailing ".0"
        box_number = (d.get('box_number') or '').strip()
        if box_number.endswith('.0') and box_number[:-2].isdigit():
            box_number = box_number[:-2]

        defaults = {
            'box_number':          box_number,
            'office_college':      (d.get('office_college')      or 'Unknown').strip(),
            'accountable_person':  (d.get('accountable_person')  or '').strip(),
            'borrower_type':       (d.get('borrower_type')       or '').strip().lower(),
            'accountable_officer': accountable_officer,
            'assigned_mr':         (d.get('assigned_mr')         or '').strip(),
            'device':              (d.get('device')              or 'Tablet').strip(),
            'ptr':                 (d.get('ptr')                 or '').strip(),
            'remarks':             (d.get('remarks')             or '').strip(),
            'issue':               (d.get('issue')               or '').strip(),
            'date_returned':       date_returned,
            'is_released':         is_released and not is_returned,
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

    # Chunked bulk DB writes
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

    # Mark returned devices in TransactionDevice
    if returned_serials:
        for i in range(0, len(returned_serials), CHUNK_SIZE):
            chunk = returned_serials[i:i + CHUNK_SIZE]
            TransactionDevice.objects.filter(
                serial_number__in=chunk,
                returned=False,
            ).update(returned=True, returned_at=now_ph)

    # Create BorrowRequest + Transaction rows for Released devices
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

    # Live broadcast
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


# ═══════════════════════════════════════════════════════════════════════════
#  EXPORT TASKS
# ═══════════════════════════════════════════════════════════════════════════

@shared_task(bind=True)
def generate_borrow_management_export(self, user_id):
    """
    Build the Borrow Management Excel file, store in cache, return a token.
    """
    from django.utils import timezone as dj_timezone
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from inventory.models import Transaction

    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).all().order_by('-borrowed_at')

    headers = [
        'Tx ID', 'Borrower Name', 'Borrower Type', 'Accountable Officer',
        'College / Office', 'Item', 'Device Serial #', 'Qty Borrowed',
        'Returned Qty', 'Borrowed On', 'Returned On',
    ]
    col_widths = [12, 24, 14, 26, 22, 20, 18, 14, 14, 20, 20]

    wb = Workbook()

    # ── Sheet 1: Transaction Details ──────────────────────────────────────
    ws_data = wb.active
    ws_data.title = 'Borrow Transactions'
    ws_data.sheet_properties.tabColor = 'FFFFFF'

    ws_data.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws_data.cell(row=1, column=1, value='Borrow Management Report')
    c.font = Font(bold=True, size=14, color='000000')
    c.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_data.row_dimensions[1].height = 30

    ws_data.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    s = ws_data.cell(row=2, column=1)
    ph_now = get_ph_time()
    s.value = f'Generated: {ph_now.strftime("%B %d, %Y %I:%M %p")}'
    s.font = Font(size=9, color='000000')
    s.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    s.alignment = Alignment(horizontal='center', vertical='center')
    ws_data.row_dimensions[2].height = 16

    fill_header = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_header = Font(bold=True, color='000000', size=11)
    border      = Border(bottom=Side(style='thin', color='CCCCCC'))
    align       = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for col, heading in enumerate(headers, start=1):
        cell = ws_data.cell(row=3, column=col, value=heading)
        cell.fill = fill_header
        cell.font = font_header
        cell.border = border
        cell.alignment = align
    ws_data.row_dimensions[3].height = 22

    summary_data = {}

    for i, tx in enumerate(transactions, start=1):
        officer = (tx.borrower.get_full_name() or '').strip() or tx.borrower.username
        college = tx.office_college or 'Unknown'
        borrower_name = tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username
        borrower_type_display = ''
        if tx.borrow_request:
            if tx.borrow_request.borrower_type == 'student':
                borrower_type_display = 'Student'
            elif tx.borrow_request.borrower_type == 'employee':
                borrower_type_display = 'Employee'

        pending_qty = tx.quantity_borrowed - tx.returned_qty

        if college not in summary_data:
            summary_data[college] = {
                'borrowed': 0,
                'returned': 0,
                'pending': 0,
                'count': 0,
                'accountable_officers': {},
            }

        summary_data[college]['borrowed'] += tx.quantity_borrowed
        summary_data[college]['returned'] += tx.returned_qty
        summary_data[college]['pending']  += pending_qty
        summary_data[college]['count']    += 1
        summary_data[college]['accountable_officers'][officer] = True

        bg_color = 'FFFFFF' if i % 2 == 0 else 'F9F9F9'
        fill_row   = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row   = Font(color='000000', size=10)
        border_row = Border(bottom=Side(style='thin', color='EEEEEE'))
        align_row  = Alignment(horizontal='center', vertical='center', wrap_text=True)

        values = [
            f'#{tx.borrow_request.transaction_id}' if tx.borrow_request else '—',
            borrower_name,
            borrower_type_display or '—',
            officer,
            college,
            tx.item.name,
            tx.serial_number or '—',
            tx.quantity_borrowed,
            tx.returned_qty,
            format_ph_time(tx.borrowed_at),
            format_ph_time(tx.returned_at) if tx.returned_at else '—',
        ]

        for col, val in enumerate(values, start=1):
            cell = ws_data.cell(row=i + 3, column=col, value=val)
            cell.fill = fill_row
            cell.font = font_row
            cell.border = border_row
            cell.alignment = align_row

    total_borrowed      = sum(d['borrowed'] for d in summary_data.values())
    total_returned      = sum(d['returned'] for d in summary_data.values())
    total_pending       = sum(d['pending']  for d in summary_data.values())
    overall_return_rate = (total_returned / total_borrowed * 100) if total_borrowed > 0 else 0

    # ── Sheet 2: Summary Report ───────────────────────────────────────────
    ws_summary = wb.create_sheet('Summary Report')
    ws_summary.sheet_properties.tabColor = 'FFFFFF'

    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_summary.cell(row=1, column=1, value='BORROW MANAGEMENT SUMMARY REPORT')
    title_cell.font = Font(bold=True, size=16, color='000000')
    title_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center')

    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    date_cell = ws_summary.cell(row=2, column=1, value=f'Report Generated: {format_ph_time(dj_timezone.now())}')
    date_cell.font = Font(size=10, color='000000')
    date_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    date_cell.alignment = Alignment(horizontal='center')

    row_num = 4
    ws_summary.cell(row=row_num, column=1, value='OVERVIEW:').font = Font(bold=True, size=12, color='000000')
    row_num += 1

    overview_text = (
        f"As of {format_ph_time(dj_timezone.now())}, there have been a total of "
        f"{transactions.count()} borrowing transactions across all colleges and offices. "
        f"A total of {total_borrowed} items have been borrowed, with {total_returned} items "
        f"successfully returned ({overall_return_rate:.1f}% return rate). "
        f"Currently, {total_pending} items are still pending return."
    )
    ws_summary.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
    ov_cell = ws_summary.cell(row=row_num, column=1, value=overview_text)
    ov_cell.alignment = Alignment(wrap_text=True)
    ov_cell.font = Font(size=11, color='000000')
    ov_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    ws_summary.row_dimensions[row_num].height = 60
    row_num += 2

    ws_summary.cell(row=row_num, column=1, value='BREAKDOWN BY COLLEGE/OFFICE:').font = Font(bold=True, size=12, color='000000')
    row_num += 1

    best_college      = None
    best_rate         = 0
    attention_colleges = []

    for college, data in sorted(summary_data.items()):
        college_return_rate = (data['returned'] / data['borrowed'] * 100) if data['borrowed'] > 0 else 0

        if college_return_rate >= 90:
            rating = 'Excellent'
        elif college_return_rate >= 70:
            rating = 'Good'
        elif college_return_rate >= 50:
            rating = 'Fair'
        else:
            rating = 'Needs Attention'
            attention_colleges.append(college)

        if college_return_rate > best_rate and data['borrowed'] > 0:
            best_rate    = college_return_rate
            best_college = college

        officers_list = ', '.join(data['accountable_officers'].keys())

        ws_summary.cell(row=row_num, column=1, value=f'{college}:').font = Font(bold=True, size=11, color='000000')
        row_num += 1

        for line in [
            f'  • Transactions: {data["count"]} | Borrowed: {data["borrowed"]} | Returned: {data["returned"]} | Pending: {data["pending"]}',
            f'  • Return Rate: {college_return_rate:.1f}% ({rating})',
            f'  • Accountable Officer(s): {officers_list}',
        ]:
            ws_summary.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
            cell = ws_summary.cell(row=row_num, column=1, value=line)
            cell.alignment = Alignment(wrap_text=True)
            cell.font = Font(size=11, color='000000')
            cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
            row_num += 1

        ws_summary.cell(row=row_num, column=1, value='')
        row_num += 1

    ws_summary.cell(row=row_num, column=1, value='KEY INSIGHTS:').font = Font(bold=True, size=12, color='000000')
    row_num += 1

    insights = []
    if best_college:
        insights.append(f'• Best Performing: {best_college} with a {best_rate:.1f}% return rate.')
    most_active = max(summary_data.items(), key=lambda x: x[1]['count']) if summary_data else (None, None)
    if most_active and most_active[0]:
        insights.append(f'• Most Active: {most_active[0]} with {most_active[1]["count"]} borrowing transaction(s).')
    if attention_colleges:
        insights.append(f'• Needs Attention: {", ".join(attention_colleges)} have return rates below 50%.')
    insights.append(f'• Overall Return Rate: {overall_return_rate:.1f}% ({total_returned} of {total_borrowed} items).')
    insights.append(f'• Outstanding Items: {total_pending} items still need to be returned.')

    for line in insights:
        ws_summary.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
        cell = ws_summary.cell(row=row_num, column=1, value=line)
        cell.alignment = Alignment(wrap_text=True)
        cell.font = Font(size=11, color='000000')
        cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        row_num += 1

    row_num += 1
    ws_summary.cell(row=row_num, column=1, value='RECOMMENDATIONS:').font = Font(bold=True, size=12, color='000000')
    row_num += 1

    recs = []
    if total_pending > 10:
        recs.append(f'• Follow up on {total_pending} outstanding items across all colleges.')
    for college in attention_colleges:
        recs.append(f'• Schedule follow-up with {college} regarding {summary_data[college]["pending"]} pending item(s).')
    if overall_return_rate < 80:
        recs.append('• Consider implementing stricter borrowing policies to improve return rates.')
    if not recs:
        recs.append('• All colleges are performing well. Continue current monitoring practices.')

    for line in recs:
        ws_summary.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
        cell = ws_summary.cell(row=row_num, column=1, value=line)
        cell.alignment = Alignment(wrap_text=True)
        cell.font = Font(size=11, color='000000')
        cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        row_num += 1

    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 50
    ws_summary.column_dimensions['C'].width = 15
    ws_summary.column_dimensions['D'].width = 15

    # ── Sheet 3: Summary Table ────────────────────────────────────────────
    ws_table = wb.create_sheet('Summary Table')
    ws_table.sheet_properties.tabColor = 'FFFFFF'

    ws_table.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    tbl_title = ws_table.cell(row=1, column=1, value='QUICK REFERENCE SUMMARY BY COLLEGE')
    tbl_title.font = Font(bold=True, size=14, color='000000')
    tbl_title.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    tbl_title.alignment = Alignment(horizontal='center')

    tbl_headers = ['College / Office', 'Accountable Officer(s)', 'Transactions', 'Borrowed', 'Returned', 'Pending', 'Return Rate']
    fill_tbl_header = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_tbl_header = Font(bold=True, color='000000', size=11)

    for col, hdr in enumerate(tbl_headers, start=1):
        cell = ws_table.cell(row=3, column=col, value=hdr)
        cell.fill = fill_tbl_header
        cell.font = font_tbl_header
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(bottom=Side(style='thin', color='CCCCCC'))

    tbl_row = 4
    for college, data in sorted(summary_data.items()):
        return_rate   = (data['returned'] / data['borrowed'] * 100) if data['borrowed'] > 0 else 0
        officers_list = ', '.join(data['accountable_officers'].keys())

        bg_color = 'FFFFFF' if tbl_row % 2 == 0 else 'F9F9F9'
        fill_r   = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_r   = Font(color='000000', size=10)

        row_values = [college, officers_list, data['count'], data['borrowed'], data['returned'], data['pending'], f'{return_rate:.1f}%']
        for col, val in enumerate(row_values, start=1):
            cell = ws_table.cell(row=tbl_row, column=col, value=val)
            cell.fill = fill_r
            cell.font = font_r
            cell.alignment = Alignment(horizontal='center', wrap_text=True)

        rate_cell = ws_table.cell(row=tbl_row, column=7)
        if return_rate >= 90:
            rate_cell.font = Font(color='00e5a0', bold=True, size=10)
        elif return_rate >= 70:
            rate_cell.font = Font(color='ffb347', bold=True, size=10)
        else:
            rate_cell.font = Font(color='ff4444', bold=True, size=10)

        tbl_row += 1

    grand_vals = ['GRAND TOTAL', '', sum(d['count'] for d in summary_data.values()),
                  total_borrowed, total_returned, total_pending, f'{overall_return_rate:.1f}%']
    for col, val in enumerate(grand_vals, start=1):
        cell = ws_table.cell(row=tbl_row, column=col, value=val)
        cell.font = Font(bold=True, color='000000', size=10)
        cell.fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')

    table_col_widths = [25, 45, 15, 12, 12, 12, 15]
    for col, width in enumerate(table_col_widths, start=1):
        ws_table.column_dimensions[get_column_letter(col)].width = width

    for col, width in enumerate(col_widths, start=1):
        ws_data.column_dimensions[get_column_letter(col)].width = width
    ws_data.freeze_panes = 'A4'

    # ── Save to cache ─────────────────────────────────────────────────────
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    file_data = buffer.getvalue()

    token = self.request.id
    cache.set(f'export_{token}', file_data, 300)
    cache.set(f'export_{token}_fn', 'borrow_management.xlsx', 300)

    return {'ok': True, 'token': token}


@shared_task(bind=True)
def generate_device_monitoring_export(self, user_id):
    """
    Build the Device Monitoring Excel file, store in cache, return a token.
    """
    import re
    from django.utils import timezone as dj_timezone
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from inventory.models import DeviceMonitor, TransactionDevice

    rows = list(DeviceMonitor.objects.all())

    def box_number_key(row):
        bn = row.box_number or ''
        match = re.search(r'(\d+)', bn)
        if match:
            return (int(match.group(1)), bn)
        return (float('inf'), bn)

    rows.sort(key=box_number_key)

    for row in rows:
        if row.date_returned:
            row.release_status = 'Returned'
        else:
            active_td = TransactionDevice.objects.filter(
                serial_number=row.serial_number,
                returned=False
            ).select_related('transaction').first()
            if active_td and active_td.transaction:
                tx = active_td.transaction
                tx_borrower = tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username
                if tx_borrower == row.accountable_person and tx.office_college == row.office_college:
                    row.release_status = 'Released'
                else:
                    row.release_status = '—'
            else:
                row.release_status = '—'

    summary_data = {}
    device_status_summary = {
        'serviceable': 0, 'non_serviceable': 0, 'sealed': 0,
        'missing': 0, 'incomplete': 0, 'released': 0, 'returned': 0,
    }
    device_type_summary = {}
    mr_stats = {}

    for row in rows:
        college = row.office_college or 'Unknown'
        assigned_mr = (row.assigned_mr or '').strip()
        if assigned_mr == '':
            assigned_mr = '—'

        if college not in summary_data:
            summary_data[college] = {
                'total_devices': 0, 'serviceable': 0, 'non_serviceable': 0,
                'sealed': 0, 'missing': 0, 'incomplete': 0,
                'released': 0, 'returned': 0, 'devices_with_issues': 0,
            }
        summary_data[college]['total_devices'] += 1
        for field in ('serviceable', 'non_serviceable', 'sealed', 'missing', 'incomplete'):
            if getattr(row, field):
                summary_data[college][field] += 1
                device_status_summary[field] += 1
                if field in ('non_serviceable', 'missing', 'incomplete'):
                    summary_data[college]['devices_with_issues'] += 1

        rs = getattr(row, 'release_status', '—')
        if rs == 'Released':
            summary_data[college]['released'] += 1
            device_status_summary['released'] += 1
        elif rs == 'Returned':
            summary_data[college]['returned'] += 1
            device_status_summary['returned'] += 1

        if assigned_mr not in mr_stats:
            mr_stats[assigned_mr] = {
                'total': 0, 'serviceable': 0, 'non_serviceable': 0, 'sealed': 0,
                'missing': 0, 'incomplete': 0, 'released': 0, 'returned': 0,
                'college_details': {},
            }
        stats = mr_stats[assigned_mr]
        stats['total'] += 1
        for field in ('serviceable', 'non_serviceable', 'sealed', 'missing', 'incomplete'):
            if getattr(row, field):
                stats[field] += 1
        if rs == 'Released':
            stats['released'] += 1
        elif rs == 'Returned':
            stats['returned'] += 1

        if college not in stats['college_details']:
            stats['college_details'][college] = {
                'total': 0, 'serviceable': 0, 'non_serviceable': 0, 'sealed': 0,
                'missing': 0, 'incomplete': 0, 'released': 0, 'returned': 0,
            }
        col_stats = stats['college_details'][college]
        col_stats['total'] += 1
        for field in ('serviceable', 'non_serviceable', 'sealed', 'missing', 'incomplete'):
            if getattr(row, field):
                col_stats[field] += 1
        if rs == 'Released':
            col_stats['released'] += 1
        elif rs == 'Returned':
            col_stats['returned'] += 1

        device = row.device or 'Tablet'
        device_type_summary[device] = device_type_summary.get(device, 0) + 1

    total_devices = len(rows)
    total_issues = (device_status_summary['non_serviceable']
                    + device_status_summary['missing']
                    + device_status_summary['incomplete'])
    health_percentage = ((total_devices - total_issues) / total_devices * 100) if total_devices > 0 else 0
    svc_pct = (device_status_summary['serviceable'] / total_devices * 100) if total_devices > 0 else 0

    wb = Workbook()

    ws_details = wb.active
    ws_details.title = 'Device Details'
    ws_details.sheet_properties.tabColor = 'FFFFFF'

    headers = [
        'Box Number', 'College / Office', 'Student', 'Borrower Type',
        'Accountable Officer', 'Assigned M.R.', 'Device', 'Serial Number',
        'Serviceable', 'Non-Serviceable', 'Sealed', 'Missing', 'Incomplete',
        'Release / Return', 'Date Returned', 'Remarks', 'Issue',
    ]
    col_widths = [15, 20, 24, 12, 24, 18, 14, 20, 14, 16, 10, 10, 12, 16, 22, 28, 28]

    ws_details.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws_details.cell(row=1, column=1, value='Device Monitoring Report')
    c.font = Font(bold=True, size=14, color='000000')
    c.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_details.row_dimensions[1].height = 30

    ws_details.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    s = ws_details.cell(row=2, column=1)
    ph_now = get_ph_time()
    s.value = f'Generated: {ph_now.strftime("%B %d, %Y %I:%M %p")}'
    s.font = Font(size=9, color='000000')
    s.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    s.alignment = Alignment(horizontal='center', vertical='center')
    ws_details.row_dimensions[2].height = 16

    fill_hdr = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_hdr = Font(bold=True, color='000000', size=11)
    bdr = Border(bottom=Side(style='thin', color='CCCCCC'))
    aln = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for col, heading in enumerate(headers, start=1):
        cell = ws_details.cell(row=3, column=col, value=heading)
        cell.fill = fill_hdr
        cell.font = font_hdr
        cell.border = bdr
        cell.alignment = aln
    ws_details.row_dimensions[3].height = 22

    for i, row in enumerate(rows, start=1):
        borrower_type_display = (
            'Student' if row.borrower_type == 'student' else
            'Employee' if row.borrower_type == 'employee' else '—'
        )
        release_status = getattr(row, 'release_status', '—')
        date_ret = format_ph_time(row.date_returned) if row.date_returned else '—'

        bg_color = 'FFFFFF' if i % 2 == 0 else 'F9F9F9'
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row = Font(color='000000', size=10)
        border_row = Border(bottom=Side(style='thin', color='EEEEEE'))
        align_row = Alignment(horizontal='center', vertical='center', wrap_text=True)

        bool_vals = [row.serviceable, row.non_serviceable, row.sealed, row.missing, row.incomplete]
        values = [
            row.box_number or '—',
            row.office_college or '—',
            row.accountable_person or '—',
            borrower_type_display,
            row.accountable_officer or '—',
            row.assigned_mr or '—',
            row.device or 'Tablet',
            row.serial_number or '—',
            '✓' if row.serviceable else '—',
            '✓' if row.non_serviceable else '—',
            '✓' if row.sealed else '—',
            '✓' if row.missing else '—',
            '✓' if row.incomplete else '—',
            release_status,
            date_ret,
            row.remarks or '—',
            row.issue or '—',
        ]

        for col, val in enumerate(values, start=1):
            cell = ws_details.cell(row=i + 3, column=col, value=val)
            cell.fill = fill_row
            cell.font = font_row
            cell.border = border_row
            cell.alignment = align_row

        for col_offset, val in enumerate(bool_vals):
            if val:
                ws_details.cell(row=i + 3, column=8 + col_offset).font = Font(color='00e5a0', bold=True, size=10)

    for col, width in enumerate(col_widths, start=1):
        ws_details.column_dimensions[get_column_letter(col)].width = width
    ws_details.freeze_panes = 'A4'

    # ── Summary Report ─────────────────────────────────────────────────────
    ws_summary = wb.create_sheet('Summary Report')
    ws_summary.sheet_properties.tabColor = 'FFFFFF'

    def write_table(ws, start_row, title, headers, data_rows, col_widths=None):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(headers))
        title_cell = ws.cell(row=start_row, column=1, value=title)
        title_cell.font = Font(bold=True, size=12, color='000000')
        title_cell.fill = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
        title_cell.alignment = Alignment(horizontal='center')
        ws.row_dimensions[start_row].height = 25
        header_row = start_row + 1

        fill_hdr2 = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
        font_hdr2 = Font(bold=True, color='000000', size=11)
        for col, hdr in enumerate(headers, start=1):
            cell = ws.cell(row=header_row, column=col, value=hdr)
            cell.fill = fill_hdr2
            cell.font = font_hdr2
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(bottom=Side(style='thin', color='888888'))
        ws.row_dimensions[header_row].height = 20

        for i, row_vals in enumerate(data_rows, start=1):
            bg = 'FFFFFF' if i % 2 == 0 else 'F9F9F9'
            fill_r = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
            font_r = Font(color='000000', size=10)
            for col, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=header_row + i, column=col, value=val)
                cell.fill = fill_r
                cell.font = font_r
                cell.alignment = Alignment(horizontal='center', wrap_text=True)
                cell.border = Border(bottom=Side(style='thin', color='EEEEEE'))
        if col_widths:
            for col, width in enumerate(col_widths, start=1):
                ws.column_dimensions[get_column_letter(col)].width = width
        return header_row + len(data_rows) + 1

    current_row = 1

    overall_data = [
        ['Total Devices', total_devices],
        ['Serviceable', f"{device_status_summary['serviceable']} ({svc_pct:.1f}%)"],
        ['Sealed', device_status_summary['sealed']],
        ['Non-Serviceable', device_status_summary['non_serviceable']],
        ['Missing', device_status_summary['missing']],
        ['Incomplete', device_status_summary['incomplete']],
        ['Devices with Issues', total_issues],
        ['Overall Device Health', f"{health_percentage:.1f}%"],
    ]
    current_row = write_table(ws_summary, current_row, '📊 OVERALL INVENTORY STATUS',
                              ['Metric', 'Value'], overall_data, [30, 20])
    current_row += 1

    if len(device_type_summary) > 1:
        dev_type_data = [[k, v] for k, v in device_type_summary.items()]
        current_row = write_table(ws_summary, current_row, '📱 DEVICE TYPE DISTRIBUTION',
                                  ['Device Type', 'Count'], dev_type_data, [25, 15])
        current_row += 1

    detail_data = []
    for mr_name in sorted(mr_stats.keys(), key=lambda x: (x == '—', x)):
        for college, col_stats in sorted(mr_stats[mr_name]['college_details'].items()):
            total_in_college = col_stats['total']
            issues = col_stats['non_serviceable'] + col_stats['missing'] + col_stats['incomplete']
            health_pct = ((total_in_college - issues) / total_in_college * 100) if total_in_college > 0 else 0
            detail_data.append([
                mr_name,
                college,
                total_in_college,
                col_stats['serviceable'],
                col_stats['non_serviceable'],
                col_stats['sealed'],
                col_stats['missing'],
                col_stats['incomplete'],
                col_stats['released'],
                col_stats['returned'],
                f"{health_pct:.1f}%",
            ])
    if detail_data:
        detail_headers = ['Assigned M.R.', 'College / Office', 'Total Devices', 'Serviceable',
                          'Non‑Svc', 'Sealed', 'Missing', 'Incomplete', 'Borrowed', 'Returned', 'Healthy %']
        detail_widths = [25, 30, 12, 12, 12, 10, 10, 12, 12, 12, 12]
        current_row = write_table(ws_summary, current_row, '🔍 DETAILED BREAKDOWN BY ASSIGNED M.R. AND COLLEGE',
                                  detail_headers, detail_data, detail_widths)
        current_row += 1

    insights_lines = [
        f"• Overall device health: {health_percentage:.1f}%",
        f"• Serviceable rate: {svc_pct:.1f}%",
    ]
    if device_status_summary['missing'] > 0:
        insights_lines.append(f"⚠️ ALERT: {device_status_summary['missing']} device(s) marked MISSING")
    if device_status_summary['non_serviceable'] > 0:
        insights_lines.append(f"🔧 {device_status_summary['non_serviceable']} device(s) need repair")
    if device_status_summary['incomplete'] > 0:
        insights_lines.append(f"📦 {device_status_summary['incomplete']} device(s) are incomplete")
    colleges_issues = [c for c, d in summary_data.items() if d['devices_with_issues'] > 0]
    if colleges_issues:
        insights_lines.append(f"⚠️ Colleges needing attention: {', '.join(colleges_issues)}")
    if device_status_summary['released'] > 0:
        insights_lines.append(f"🔄 {device_status_summary['released']} device(s) currently borrowed")
    insight_text = '\n'.join(insights_lines)
    merge_cols = len(detail_headers) if detail_headers else 11
    ws_summary.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=merge_cols)
    cell = ws_summary.cell(row=current_row, column=1, value='💡 KEY INSIGHTS\n' + insight_text)
    cell.font = Font(bold=True, size=11, color='000000')
    cell.fill = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    cell.alignment = Alignment(wrap_text=True, horizontal='left')
    ws_summary.row_dimensions[current_row].height = 30 + 15 * len(insights_lines)
    current_row += 2

    recs = []
    if device_status_summary['missing'] > 0:
        recs.append(f"🔴 Conduct physical inventory for {device_status_summary['missing']} missing device(s)")
    if device_status_summary['non_serviceable'] > 0:
        recs.append(f"🔧 Schedule repair for {device_status_summary['non_serviceable']} non‑serviceable devices")
    if device_status_summary['incomplete'] > 0:
        recs.append(f"📋 Audit {device_status_summary['incomplete']} incomplete devices")
    for college in colleges_issues:
        recs.append(f"📞 Follow up with {college} ({summary_data[college]['devices_with_issues']} device(s) with issues)")
    if not recs:
        recs.append("✅ All devices in good condition. Continue regular monitoring.")
    rec_text = '\n'.join(recs)
    ws_summary.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=merge_cols)
    cell = ws_summary.cell(row=current_row, column=1, value='🎯 RECOMMENDATIONS\n' + rec_text)
    cell.font = Font(bold=True, size=11, color='000000')
    cell.fill = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    cell.alignment = Alignment(wrap_text=True, horizontal='left')
    ws_summary.row_dimensions[current_row].height = 30 + 20 * len(recs)
    current_row += 2

    # ── Save to cache ─────────────────────────────────────────────────────
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    file_data = buffer.getvalue()

    token = self.request.id
    cache.set(f'export_{token}', file_data, 300)
    cache.set(f'export_{token}_fn', 'device_monitoring.xlsx', 300)

    return {'ok': True, 'token': token}