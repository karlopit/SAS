"""
Patch notes — only the borrow_management view changes meaningfully.
The full views.py is reproduced here so it can be dropped in as a replacement.
"""
import io
import json
import random
from django.db.models import Sum, F, ExpressionWrapper, IntegerField
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.http import JsonResponse, HttpResponse
from django.utils import timezone
from django.views.decorators.http import require_POST
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .models import Item, Transaction, BorrowRequest, DeviceMonitor
from .forms import ItemForm, StaffBorrowForm, TransactionConditionForm, BorrowRequestForm
from .decorators import no_cache
from django.contrib import messages


def _broadcasts():
    from inventory import broadcasts as b
    return b


# ─────────────────────────────────────────────────────────────────────────────
#  Public / unauthenticated
# ─────────────────────────────────────────────────────────────────────────────

def welcome(request):
    if request.user.is_authenticated:
        return redirect('index')

    borrow_form     = BorrowRequestForm()
    borrow_success  = None
    generated_tx_id = str(random.randint(10000, 99999))

    if 'borrow_success' in request.session:
        borrow_success  = request.session.pop('borrow_success')
        generated_tx_id = str(random.randint(10000, 99999))

    if request.method == 'POST' and request.POST.get('action') == 'borrow_request':
        borrow_form = BorrowRequestForm(request.POST)
        if borrow_form.is_valid():
            req   = borrow_form.save(commit=False)
            tx_id = request.POST.get('transaction_id', str(random.randint(10000, 99999)))
            while BorrowRequest.objects.filter(transaction_id=tx_id).exists():
                tx_id = str(random.randint(10000, 99999))
            req.transaction_id = tx_id
            req.save()

            request.session['borrow_success'] = req.transaction_id
            b = _broadcasts()
            b.broadcast_borrow_requests()
            b.broadcast_dashboard()
            return redirect('welcome')

    return render(request, 'inventory/welcome.html', {
        'borrow_form':     borrow_form,
        'borrow_success':  borrow_success,
        'generated_tx_id': generated_tx_id,
        'available_items': Item.objects.filter(available_quantity__gt=0),
    })


# ─────────────────────────────────────────────────────────────────────────────
#  Dashboard
# ─────────────────────────────────────────────────────────────────────────────

@login_required
@no_cache
def index(request):
    pending_count  = BorrowRequest.objects.filter(status='pending').count()
    items          = Item.objects.all()
    active_borrows = Transaction.objects.filter(status='borrowed').count()
    total_returns  = Transaction.objects.filter(status='returned').count()
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

    return render(request, 'inventory/index.html', {
        'items':          items,
        'active_borrows': active_borrows,
        'total_returns':  total_returns,
        'pending_count':  pending_count,
        'available_qty':  available_qty,
        'borrowed_qty':   borrowed_qty,
        'dm_offices':     json.dumps(offices),
        'dm_serviceable': json.dumps([monitors.filter(office_college=o, serviceable=True).count()     for o in offices]),
        'dm_non_service': json.dumps([monitors.filter(office_college=o, non_serviceable=True).count() for o in offices]),
        'dm_sealed':      json.dumps([monitors.filter(office_college=o, sealed=True).count()          for o in offices]),
        'dm_missing':     json.dumps([monitors.filter(office_college=o, missing=True).count()         for o in offices]),
        'dm_incomplete':  json.dumps([monitors.filter(office_college=o, incomplete=True).count()      for o in offices]),
    })


# ─────────────────────────────────────────────────────────────────────────────
#  AJAX poll endpoints
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def ajax_dashboard_data(request):
    from inventory.consumers import _build_dashboard_payload
    return JsonResponse(_build_dashboard_payload())


@login_required
def ajax_borrow_management_data(request):
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)
    from inventory.consumers import _build_borrow_management_payload
    return JsonResponse(_build_borrow_management_payload())


@login_required
def ajax_borrow_requests_data(request):
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)
    from inventory.consumers import _build_borrow_requests_payload
    return JsonResponse(_build_borrow_requests_payload())


@login_required
def ajax_device_monitoring_data(request):
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)
    from inventory.consumers import _build_device_monitoring_payload
    return JsonResponse(_build_device_monitoring_payload())


# ─────────────────────────────────────────────────────────────────────────────
#  Item management
# ─────────────────────────────────────────────────────────────────────────────

@login_required
@no_cache
def add_item(request):
    if request.user.role != 'admin':
        raise PermissionDenied
    form = ItemForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        item = form.save(commit=False)
        item.available_quantity = item.quantity
        item.save()
        b = _broadcasts()
        b.broadcast_dashboard()
        b.broadcast_borrow_management()
        return redirect('index')
    return render(request, 'inventory/add_item.html', {'form': form})


@login_required
def edit_item(request, item_id):
    if request.user.role != 'admin':
        messages.error(request, 'You do not have permission to edit items.')
        return redirect('index')

    item = get_object_or_404(Item, id=item_id)

    if request.method == 'POST':
        new_quantity = request.POST.get('available_quantity')
        if new_quantity is not None:
            try:
                new_quantity = int(new_quantity)
                if new_quantity >= 0:
                    item.available_quantity = new_quantity
                    item.save()
                    messages.success(request, f'Updated {item.name} to {item.available_quantity} units.')
                    b = _broadcasts()
                    b.broadcast_dashboard()
                    b.broadcast_borrow_management()
                else:
                    messages.error(request, 'Quantity cannot be negative.')
            except ValueError:
                messages.error(request, 'Invalid quantity value.')
        else:
            messages.error(request, 'No quantity provided.')
        return redirect('index')

    return redirect('index')


# ─────────────────────────────────────────────────────────────────────────────
#  Borrow requests
# ─────────────────────────────────────────────────────────────────────────────

@login_required
@no_cache
def borrow_requests(request):
    if request.user.role != 'staff':
        raise PermissionDenied
    pending       = BorrowRequest.objects.filter(status='pending').order_by('-created_at')
    pending_count = pending.count()
    return render(request, 'inventory/borrow_requests.html', {
        'pending':       pending,
        'pending_count': pending_count,
    })


@login_required
@no_cache
def borrow_management(request):
    if request.user.role != 'staff':
        raise PermissionDenied

    items        = Item.objects.all()
    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).order_by('-borrowed_at')[:50]
    pending_count = BorrowRequest.objects.filter(status='pending').count()

    return render(request, 'inventory/borrow_management.html', {
        'items':         items,
        'transactions':  transactions,
        'pending_count': pending_count,
    })


@login_required
@no_cache
def device_monitoring(request):
    if request.user.role != 'staff':
        raise PermissionDenied
    rows          = DeviceMonitor.objects.all()
    pending_count = BorrowRequest.objects.filter(status='pending').count()
    return render(request, 'inventory/device_monitoring.html', {
        'rows':          rows,
        'pending_count': pending_count,
    })


@login_required
@require_POST
def device_monitoring_save(request):
    if request.user.role != 'staff':
        raise PermissionDenied

    ids              = request.POST.getlist('row_id')
    box_numbers      = request.POST.getlist('box_number')
    offices          = request.POST.getlist('office_college')
    accountables     = request.POST.getlist('accountable_person')
    accountable_officers = request.POST.getlist('accountable_officer')
    devices          = request.POST.getlist('device')
    serials          = request.POST.getlist('serial_number')
    serviceables     = request.POST.getlist('serviceable')
    non_serviceables = request.POST.getlist('non_serviceable')
    sealeds          = request.POST.getlist('sealed')
    missings         = request.POST.getlist('missing')
    incompletes      = request.POST.getlist('incomplete')

    for i, row_id in enumerate(ids):
        def get(lst, idx=i):
            return lst[idx] if idx < len(lst) else ''

        fields = dict(
            box_number          = get(box_numbers),
            office_college      = get(offices),
            accountable_person  = get(accountables),
            accountable_officer = get(accountable_officers),
            device              = get(devices) or 'Tablet',
            serial_number       = get(serials),
            serviceable         = get(serviceables)     == 'on',
            non_serviceable     = get(non_serviceables) == 'on',
            sealed              = get(sealeds)          == 'on',
            missing             = get(missings)         == 'on',
            incomplete          = get(incompletes)      == 'on',
        )

        if row_id == 'new':
            DeviceMonitor.objects.create(**fields)
        else:
            try:
                obj = DeviceMonitor.objects.get(pk=int(row_id))
                for attr, val in fields.items():
                    setattr(obj, attr, val)
                obj.save()
            except DeviceMonitor.DoesNotExist:
                pass

    b = _broadcasts()
    b.broadcast_device_monitoring()
    b.broadcast_dashboard()
    return redirect('device_monitoring')


@login_required
@require_POST
def device_monitoring_delete(request, row_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    obj = get_object_or_404(DeviceMonitor, pk=row_id)
    obj.delete()
    b = _broadcasts()
    b.broadcast_device_monitoring()
    b.broadcast_dashboard()
    return redirect('device_monitoring')


# ─────────────────────────────────────────────────────────────────────────────
#  Staff borrow confirmation / decline
# ─────────────────────────────────────────────────────────────────────────────

@login_required
@no_cache
def staff_confirm_borrow(request, request_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    borrow_req = get_object_or_404(BorrowRequest, id=request_id, status='pending')

    if request.method == 'POST':
        form = StaffBorrowForm(request.POST)
        if form.is_valid():
            serial_numbers = form.cleaned_data['serial_numbers']  # List of serial numbers
            box_numbers = form.cleaned_data['box_numbers']  # List of box numbers
            quantity = form.cleaned_data['quantity_borrowed']
            
            # Create ONE Transaction record (for borrow management)
            transaction = form.save(commit=False)
            transaction.borrower = request.user
            transaction.borrow_request = borrow_req
            transaction.office_college = borrow_req.office_college
            transaction.status = 'borrowed'
            transaction.serial_number = ', '.join(serial_numbers)  # Store all serials as comma-separated
            transaction.item.available_quantity -= quantity
            transaction.item.save()
            transaction.save()
            
            # Update borrow request status
            borrow_req.status = 'accepted'
            borrow_req.save()
            
            # ── Create MULTIPLE Device Monitor records (one per serial/box pair) ──
            accountable_officer = request.user.get_full_name() or request.user.username
            
            device_monitors = []
            for i, serial in enumerate(serial_numbers):
                # Get the corresponding box number for this serial number
                box_number = box_numbers[i] if i < len(box_numbers) else ''
                
                # REMOVED: display_id field
                device_monitor = DeviceMonitor(
                    box_number=box_number,
                    office_college=borrow_req.office_college,
                    accountable_person=borrow_req.borrower_name,
                    accountable_officer=accountable_officer,
                    device=transaction.item.name,
                    serial_number=serial,
                    serviceable=True,
                    non_serviceable=False,
                    sealed=False,
                    missing=False,
                    incomplete=False,
                )
                device_monitors.append(device_monitor)
            
            # Bulk create all device monitor records
            DeviceMonitor.objects.bulk_create(device_monitors)
            # ─────────────────────────────────────────────────────────────

            b = _broadcasts()
            b.broadcast_all()
            return redirect('index')
    else:
        form = StaffBorrowForm(initial={
            'quantity_borrowed': borrow_req.quantity,
            'office_college': borrow_req.office_college,
        })

    return render(request, 'inventory/staff_confirm_borrow.html', {
        'form': form,
        'borrow_req': borrow_req,
    })


@login_required
@no_cache
def decline_request(request, request_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    borrow_req = get_object_or_404(BorrowRequest, id=request_id, status='pending')
    if request.method == 'POST':
        borrow_req.status = 'declined'
        borrow_req.save()
        b = _broadcasts()
        b.broadcast_borrow_requests()
        b.broadcast_dashboard()
    return redirect('borrow_requests')


# ─────────────────────────────────────────────────────────────────────────────
#  Return / condition
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def return_item(request, transaction_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    transaction = get_object_or_404(Transaction, id=transaction_id)
    if request.method == 'POST' and transaction.status != 'returned':
        transaction.status      = 'returned'
        transaction.returned_at = timezone.now()
        transaction.save()
        b = _broadcasts()
        b.broadcast_borrow_management()
        b.broadcast_dashboard()
        return redirect('borrow_management')
    return render(request, 'inventory/return_item.html', {'transaction': transaction})


@login_required
@no_cache
def update_condition(request, transaction_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    tx = get_object_or_404(Transaction, id=transaction_id)
    if request.method == 'POST':
        form = TransactionConditionForm(request.POST, instance=tx)
        if form.is_valid():
            form.save()
            b = _broadcasts()
            b.broadcast_borrow_management()
    return redirect('borrow_management')


@login_required
@require_POST
def update_returned_qty(request, transaction_id):
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)

    tx = get_object_or_404(Transaction, id=transaction_id)

    try:
        new_returned = int(request.POST.get('returned_qty', 0))
    except (ValueError, TypeError):
        return JsonResponse({'error': 'Invalid value'}, status=400)

    new_returned = max(0, min(new_returned, tx.quantity_borrowed))
    delta        = new_returned - tx.returned_qty

    if delta != 0:
        tx.item.available_quantity = max(0, tx.item.available_quantity + delta)
        tx.item.save()

    tx.returned_qty = new_returned
    tx.returned_at  = timezone.now() if new_returned > 0 else None
    tx.status       = 'returned' if new_returned >= tx.quantity_borrowed else 'borrowed'
    tx.save()

    b = _broadcasts()
    b.broadcast_borrow_management()
    b.broadcast_dashboard()

    items         = Item.objects.all()
    available_qty = sum(i.available_quantity for i in items)
    agg           = Transaction.objects.annotate(
        still_out=ExpressionWrapper(
            F('quantity_borrowed') - F('returned_qty'),
            output_field=IntegerField()
        )
    ).aggregate(total=Sum('still_out'))
    borrowed_qty = max(0, agg['total'] or 0)

    return JsonResponse({
        'ok':             True,
        'returned_qty':   tx.returned_qty,
        'status':         tx.status,
        'returned_at':    tx.returned_at.strftime('%b %d, %Y %H:%M') if tx.returned_at else None,
        'fully_returned': tx.returned_qty >= tx.quantity_borrowed,
        'pie': {'available': available_qty, 'borrowed': borrowed_qty},
    })


# ─────────────────────────────────────────────────────────────────────────────
#  Excel exports
# ─────────────────────────────────────────────────────────────────────────────

def _xl_title(ws, text, col_count):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
    c = ws.cell(row=1, column=1)
    c.value = text; c.font = Font(bold=True, size=14, color='00E5A0')
    c.fill = PatternFill(start_color='0E0F13', end_color='0E0F13', fill_type='solid')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_count)
    s = ws.cell(row=2, column=1)
    s.value = f'Generated: {timezone.now().strftime("%B %d, %Y  %H:%M")}'
    s.font = Font(size=9, color='6B7080')
    s.fill = PatternFill(start_color='0E0F13', end_color='0E0F13', fill_type='solid')
    s.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 16


def _xl_header(ws, row_num, headers):
    fill = PatternFill(start_color='1E2029', end_color='1E2029', fill_type='solid')
    font = Font(bold=True, color='00E5A0', size=11)
    border = Border(bottom=Side(style='thin', color='2A2D3A'))
    align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col, heading in enumerate(headers, start=1):
        c = ws.cell(row=row_num, column=col, value=heading)
        c.fill = fill; c.font = font; c.border = border; c.alignment = align
    ws.row_dimensions[row_num].height = 22


def _xl_row(ws, row_num, values, even=False):
    bg     = '1A1C24' if even else '16181F'
    fill   = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    font   = Font(color='E8EAF0', size=10)
    border = Border(bottom=Side(style='thin', color='2A2D3A'))
    align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col, val in enumerate(values, start=1):
        c = ws.cell(row=row_num, column=col, value=val)
        c.fill = fill; c.font = font; c.border = border; c.alignment = align


def _xl_response(wb, filename_prefix):
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    filename = f'{filename_prefix}_{timezone.now().strftime("%Y%m%d_%H%M")}.xlsx'
    resp = HttpResponse(buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp


@login_required
def export_borrow_management(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied

    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).all().order_by('-borrowed_at')

    # Updated headers to include Accountable Officer
    headers    = [
        'Tx ID', 'Borrower Name', 'Accountable Officer', 'College / Office',
        'Item', 'Device Serial #', 'Qty Borrowed', 'Returned Qty',
        'Borrowed On', 'Returned On',
    ]
    col_widths = [12, 24, 22, 20, 20, 18, 14, 14, 16, 20]

    wb = Workbook(); ws = wb.active; ws.title = 'Borrow Management'
    ws.sheet_properties.tabColor = '00E5A0'
    _xl_title(ws, 'Borrow Management Report', len(headers))
    _xl_header(ws, 3, headers)

    for i, tx in enumerate(transactions, start=1):
        officer = tx.borrower.get_full_name() or tx.borrower.username
        _xl_row(ws, i + 3, [
            f'#{tx.borrow_request.transaction_id}' if tx.borrow_request else '—',
            tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username,
            officer,
            tx.office_college or '—',
            tx.item.name,
            tx.serial_number or '—',
            tx.quantity_borrowed,
            tx.returned_qty,
            tx.borrowed_at.strftime('%b %d, %Y'),
            tx.returned_at.strftime('%b %d, %Y  %H:%M') if tx.returned_at else '—',
        ], even=(i % 2 == 0))

    for col, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = 'A4'
    return _xl_response(wb, 'borrow_management')


@login_required
def export_device_monitoring(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied

    rows       = DeviceMonitor.objects.all().order_by('id')
    # Removed 'ID' column header since display_id is gone
    headers    = ['Box Number', 'College / Office', 'Accountable Person', 'Accountable Officer', 'Device', 'Serial Number',
                  'Serviceable', 'Non-Serviceable', 'Sealed', 'Missing', 'Incomplete']
    col_widths = [15, 20, 24, 24, 14, 20, 14, 16, 10, 10, 12]

    wb = Workbook(); ws = wb.active; ws.title = 'Device Monitoring'
    ws.sheet_properties.tabColor = '00E5A0'
    _xl_title(ws, 'Device Monitoring Report', len(headers))
    _xl_header(ws, 3, headers)
    font_yes = Font(bold=True, color='00E5A0', size=10)
    font_no  = Font(color='6B7080', size=10)

    for i, row in enumerate(rows, start=1):
        bool_vals = [row.serviceable, row.non_serviceable, row.sealed, row.missing, row.incomplete]
        _xl_row(ws, i + 3, [
            row.box_number or '—',
            row.office_college or '—',
            row.accountable_person or '—',
            row.accountable_officer or '—',
            row.device or 'Tablet',
            row.serial_number or '—',
            '✓' if row.serviceable     else '—',
            '✓' if row.non_serviceable else '—',
            '✓' if row.sealed          else '—',
            '✓' if row.missing         else '—',
            '✓' if row.incomplete      else '—',
        ], even=(i % 2 == 0))
        for col_offset, val in enumerate(bool_vals):
            ws.cell(row=i + 3, column=7 + col_offset).font = font_yes if val else font_no

    for col, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = 'A4'
    return _xl_response(wb, 'device_monitoring')