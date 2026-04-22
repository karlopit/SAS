import io
import json
import random
import pytz
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
from .models import Item, Transaction, BorrowRequest, DeviceMonitor, TransactionDevice
from .models import Item, Transaction, BorrowRequest, DeviceMonitor
from .forms import ItemForm, StaffBorrowForm, TransactionConditionForm, BorrowRequestForm
from .decorators import no_cache
from django.contrib import messages

# Get Philippine timezone
PH_TZ = pytz.timezone('Asia/Manila')

def get_ph_time(dt=None):
    """Return current time or converted datetime in Philippine timezone"""
    if dt is None:
        dt = timezone.now()
    if timezone.is_naive(dt):
        dt = timezone.make_aware(dt, timezone.utc)
    return dt.astimezone(PH_TZ)

def format_ph_time(dt):
    """Format datetime to Philippine time 12-hour format"""
    if not dt:
        return None
    ph_dt = get_ph_time(dt)
    return ph_dt.strftime('%b %d, %Y %I:%M %p')

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

@login_required
def transaction_devices_json(request, transaction_id):
    """
    Returns JSON list of all TransactionDevice rows for a transaction.
    Falls back to parsing Transaction.serial_number if no TransactionDevice rows
    exist yet (for transactions created before this feature was added).
    """
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)
 
    tx = get_object_or_404(Transaction, id=transaction_id)
    devices = list(tx.devices.all())
 
    # ── Fallback: build virtual device list from comma-separated serial_number ──
    if not devices and tx.serial_number:
        serials = [s.strip() for s in tx.serial_number.split(',') if s.strip()]
        # Try to find matching DeviceMonitor rows to get box numbers
        dm_map = {}
        for dm in DeviceMonitor.objects.filter(serial_number__in=serials):
            dm_map[dm.serial_number] = dm.box_number
 
        data = []
        for sn in serials:
            data.append({
                'id':            None,
                'serial_number': sn,
                'box_number':    dm_map.get(sn, '—'),
                'returned':      False,
                'returned_at':   None,
            })
        return JsonResponse({'devices': data})
 
    data = []
    for d in devices:
        data.append({
            'id':            d.id,
            'serial_number': d.serial_number,
            'box_number':    d.box_number or '—',
            'returned':      d.returned,
            'returned_at':   format_ph_time(d.returned_at),
        })
    return JsonResponse({'devices': data})


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

    items = Item.objects.all()
    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).order_by('-borrowed_at')[:50]
    
    # Add formatted time display for each transaction
    for tx in transactions:
        if tx.returned_at:
            tx.returned_at_display = format_ph_time(tx.returned_at)
        else:
            tx.returned_at_display = '—'
        if tx.borrowed_at:
            tx.borrowed_at_display = format_ph_time(tx.borrowed_at)
        else:
            tx.borrowed_at_display = '—'
    
    pending_count = BorrowRequest.objects.filter(status='pending').count()

    return render(request, 'inventory/borrow_management.html', {
        'items': items,
        'transactions': transactions,
        'pending_count': pending_count,
    })


@login_required
@no_cache
def device_monitoring(request):
    if request.user.role != 'staff':
        raise PermissionDenied
 
    rows = list(DeviceMonitor.objects.all())
    
    # Annotate each DeviceMonitor row
    for row in rows:
        # Use the device's own date_returned field
        if row.date_returned:
            row.release_status = 'Returned'
            row.date_returned_display = format_ph_time(row.date_returned)
        else:
            # Check if there's an active transaction for this device
            active_td = TransactionDevice.objects.filter(
                serial_number=row.serial_number,
                returned=False
            ).select_related('transaction').first()
            
            if active_td and active_td.transaction:
                tx = active_td.transaction
                tx_borrower = tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username
                
                if tx_borrower == row.accountable_person and tx.office_college == row.office_college:
                    row.release_status = 'Released'
                    row.date_returned_display = '—'
                else:
                    row.release_status = '—'
                    row.date_returned_display = '—'
            else:
                row.release_status = '—'
                row.date_returned_display = '—'
    
    pending_count = BorrowRequest.objects.filter(status='pending').count()
 
    return render(request, 'inventory/device_monitoring.html', {
        'rows': rows,
        'pending_count': pending_count,
    })


@login_required
@require_POST
def device_monitoring_save(request):
    if request.user.role != 'staff':
        raise PermissionDenied
 
    ids                  = request.POST.getlist('row_id')
    box_numbers          = request.POST.getlist('box_number')
    offices              = request.POST.getlist('office_college')
    accountables         = request.POST.getlist('accountable_person')
    borrower_types       = request.POST.getlist('borrower_type')
    accountable_officers = request.POST.getlist('accountable_officer')
    devices              = request.POST.getlist('device')
    serials              = request.POST.getlist('serial_number')
    serviceables         = request.POST.getlist('serviceable')
    non_serviceables     = request.POST.getlist('non_serviceable')
    sealeds              = request.POST.getlist('sealed')
    missings             = request.POST.getlist('missing')
    incompletes          = request.POST.getlist('incomplete')
    remarks_list         = request.POST.getlist('remarks')
    issue_list           = request.POST.getlist('issue')
 
    for i, row_id in enumerate(ids):
        def get(lst, idx=i):
            return lst[idx] if idx < len(lst) else ''
 
        fields = dict(
            box_number          = get(box_numbers),
            office_college      = get(offices),
            accountable_person  = get(accountables),
            borrower_type       = get(borrower_types),
            accountable_officer = get(accountable_officers),
            device              = get(devices) or 'Tablet',
            serial_number       = get(serials),
            serviceable         = get(serviceables)     == 'on',
            non_serviceable     = get(non_serviceables) == 'on',
            sealed              = get(sealeds)          == 'on',
            missing             = get(missings)         == 'on',
            incomplete          = get(incompletes)      == 'on',
            remarks             = get(remarks_list),
            issue               = get(issue_list),
        )
 
        if row_id == 'new':
            DeviceMonitor.objects.create(**fields)
        else:
            try:
                obj = DeviceMonitor.objects.get(pk=int(row_id))
                existing_date_returned = obj.date_returned
                for attr, val in fields.items():
                    setattr(obj, attr, val)
                obj.date_returned = existing_date_returned
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
            serial_numbers = form.cleaned_data['serial_numbers']
            box_numbers    = form.cleaned_data['box_numbers']
            quantity       = form.cleaned_data['quantity_borrowed']
 
            transaction = form.save(commit=False)
            transaction.borrower       = request.user
            transaction.borrow_request = borrow_req
            transaction.office_college = borrow_req.office_college
            transaction.status         = 'borrowed'
            transaction.serial_number  = ', '.join(serial_numbers)
            transaction.item.available_quantity -= quantity
            transaction.item.save()
            transaction.save()
 
            borrow_req.status = 'accepted'
            borrow_req.save()
 
            accountable_officer = request.user.get_full_name() or request.user.username
 
            device_monitors = []
            for i, serial in enumerate(serial_numbers):
                box = box_numbers[i] if i < len(box_numbers) else ''
 
                TransactionDevice.objects.create(
                    transaction=transaction,
                    serial_number=serial,
                    box_number=box,
                    returned=False,
                    returned_at=None,
                )
 
                device_monitors.append(DeviceMonitor(
                    box_number=box,
                    office_college=borrow_req.office_college,
                    accountable_person=borrow_req.borrower_name,
                    borrower_type=borrow_req.borrower_type,
                    accountable_officer=accountable_officer,
                    device=transaction.item.name,
                    serial_number=serial,
                    serviceable=True,
                    non_serviceable=False,
                    sealed=False,
                    missing=False,
                    incomplete=False,
                ))
 
            DeviceMonitor.objects.bulk_create(device_monitors)
 
            b = _broadcasts()
            b.broadcast_all()
            return redirect('index')
    else:
        form = StaffBorrowForm(initial={
            'quantity_borrowed': borrow_req.quantity,
            'office_college':    borrow_req.office_college,
        })
 
    return render(request, 'inventory/staff_confirm_borrow.html', {
        'form':       form,
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
        transaction.status = 'returned'
        transaction.returned_at = get_ph_time()
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
    tx.returned_at = get_ph_time() if new_returned > 0 else None
    tx.status = 'returned' if new_returned >= tx.quantity_borrowed else 'borrowed'
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
        'returned_at':    format_ph_time(tx.returned_at),
        'fully_returned': tx.returned_qty >= tx.quantity_borrowed,
        'pie': {'available': available_qty, 'borrowed': borrowed_qty},
    })

@login_required
@require_POST
def return_devices(request, transaction_id):
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)
 
    tx = get_object_or_404(Transaction, id=transaction_id)
    
    try:
        body = json.loads(request.body)
        device_ids = body.get('device_ids', [])
        serials = body.get('serials', [])
    except (json.JSONDecodeError, AttributeError) as e:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)
 
    # Get current Philippine time
    now_ph = get_ph_time()
    returned_serials = []
 
    # Update TransactionDevice records
    if device_ids:
        real_ids = [d for d in device_ids if d is not None]
        if real_ids:
            updated_devices = TransactionDevice.objects.filter(
                id__in=real_ids,
                transaction=tx,
                returned=False,
            )
            returned_serials = list(updated_devices.values_list('serial_number', flat=True))
            updated_devices.update(returned=True, returned_at=now_ph)
    elif serials:
        for sn in serials:
            td = tx.devices.filter(serial_number=sn, returned=False).first()
            if td:
                td.returned = True
                td.returned_at = now_ph
                td.save()
                returned_serials.append(sn)
 
    # Update DeviceMonitor records
    if returned_serials:
        if tx.borrow_request:
            borrower_name = tx.borrow_request.borrower_name
            office = tx.borrow_request.office_college
        else:
            borrower_name = tx.borrower.get_full_name() or tx.borrower.username
            office = tx.office_college
        
        for serial in returned_serials:
            DeviceMonitor.objects.filter(
                serial_number=serial,
                accountable_person=borrower_name,
                office_college=office,
                date_returned__isnull=True
            ).update(date_returned=now_ph)
    
    # Recalculate returned_qty
    returned_count = tx.devices.filter(returned=True).count()
 
    if not tx.devices.exists():
        returned_count = tx.returned_qty + len(serials)
 
    returned_count = min(returned_count, tx.quantity_borrowed)
 
    delta = returned_count - tx.returned_qty
    if delta > 0:
        tx.item.available_quantity = tx.item.available_quantity + delta
        tx.item.save()
 
    tx.returned_qty = returned_count
    tx.returned_at = now_ph if returned_count > 0 else tx.returned_at
    tx.status = 'returned' if returned_count >= tx.quantity_borrowed else 'borrowed'
    tx.save()
 
    b = _broadcasts()
    b.broadcast_borrow_management()
    b.broadcast_dashboard()
    b.broadcast_device_monitoring()
 
    return JsonResponse({
        'ok': True,
        'returned_qty': tx.returned_qty,
        'status': tx.status,
        'fully_returned': tx.returned_qty >= tx.quantity_borrowed,
        'returned_at': format_ph_time(tx.returned_at),
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
    ph_now = get_ph_time()
    s.value = f'Generated: {ph_now.strftime("%B %d, %Y %I:%M %p")}'
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
    ph_now = get_ph_time()
    filename = f'{filename_prefix}_{ph_now.strftime("%Y%m%d_%H%M")}.xlsx'
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

    headers = [
        'Tx ID', 'Borrower Name', 'Accountable Officer', 'College / Office',
        'Item', 'Device Serial #', 'Qty Borrowed', 'Returned Qty',
        'Borrowed On', 'Returned On',
    ]
    col_widths = [12, 24, 22, 20, 20, 18, 14, 14, 16, 20]

    wb = Workbook()
    
    # ========== SHEET 1: Transaction Details ==========
    ws_data = wb.active
    ws_data.title = 'Borrow Transactions'
    ws_data.sheet_properties.tabColor = 'FFFFFF'
    
    # Custom title with white background, black text
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
    
    # Header row
    fill_header = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_header = Font(bold=True, color='000000', size=11)
    border = Border(bottom=Side(style='thin', color='CCCCCC'))
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for col, heading in enumerate(headers, start=1):
        cell = ws_data.cell(row=3, column=col, value=heading)
        cell.fill = fill_header
        cell.font = font_header
        cell.border = border
        cell.alignment = align
    ws_data.row_dimensions[3].height = 22

    # Collect summary data by college/office
    summary_data = {}
    
    for i, tx in enumerate(transactions, start=1):
        officer = tx.borrower.get_full_name() or tx.borrower.username
        college = tx.office_college or 'Unknown'
        
        pending_qty = tx.quantity_borrowed - tx.returned_qty
        
        if college not in summary_data:
            summary_data[college] = {
                'borrowed': 0,
                'returned': 0,
                'pending': 0,
                'count': 0,
                'accountable_officers': {}
            }
        
        summary_data[college]['borrowed'] += tx.quantity_borrowed
        summary_data[college]['returned'] += tx.returned_qty
        summary_data[college]['pending'] += pending_qty
        summary_data[college]['count'] += 1
        summary_data[college]['accountable_officers'][officer] = True
        
        # Write transaction row
        bg_color = 'FFFFFF' if i % 2 == 0 else 'F9F9F9'
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row = Font(color='000000', size=10)
        border_row = Border(bottom=Side(style='thin', color='EEEEEE'))
        align_row = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        values = [
            f'#{tx.borrow_request.transaction_id}' if tx.borrow_request else '—',
            tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username,
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

    # Calculate overall totals
    total_borrowed = sum(data['borrowed'] for data in summary_data.values())
    total_returned = sum(data['returned'] for data in summary_data.values())
    total_pending = sum(data['pending'] for data in summary_data.values())
    overall_return_rate = (total_returned / total_borrowed * 100) if total_borrowed > 0 else 0

    # ========== SHEET 2: Summary Report ==========
    ws_summary = wb.create_sheet('Summary Report')
    ws_summary.sheet_properties.tabColor = 'FFFFFF'
    
    # Title
    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_summary.cell(row=1, column=1, value='BORROW MANAGEMENT SUMMARY REPORT')
    title_cell.font = Font(bold=True, size=16, color='000000')
    title_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center')
    
    # Generation date
    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    date_cell = ws_summary.cell(row=2, column=1, value=f'Report Generated: {format_ph_time(timezone.now())}')
    date_cell.font = Font(size=10, color='000000')
    date_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    date_cell.alignment = Alignment(horizontal='center')
    
    row_num = 4
    
    # Overview
    ws_summary.cell(row=row_num, column=1, value="OVERVIEW:").font = Font(bold=True, size=12, color='000000')
    row_num += 1
    
    overview_text = f"As of {format_ph_time(timezone.now())}, there have been a total of {transactions.count()} borrowing transactions across all colleges and offices. A total of {total_borrowed} items have been borrowed, with {total_returned} items successfully returned ({overall_return_rate:.1f}% return rate). Currently, {total_pending} items are still pending return."
    ws_summary.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=4)
    ws_summary.cell(row=row_num, column=1, value=overview_text).alignment = Alignment(wrap_text=True)
    row_num += 2
    
    # Breakdown by College
    ws_summary.cell(row=row_num, column=1, value="BREAKDOWN BY COLLEGE/OFFICE:").font = Font(bold=True, size=12, color='000000')
    row_num += 1
    
    best_college = None
    best_rate = 0
    attention_colleges = []
    
    for college, data in sorted(summary_data.items()):
        college_return_rate = (data['returned'] / data['borrowed'] * 100) if data['borrowed'] > 0 else 0
        
        if college_return_rate >= 90:
            rating = "Excellent"
        elif college_return_rate >= 70:
            rating = "Good"
        elif college_return_rate >= 50:
            rating = "Fair"
        else:
            rating = "Needs Attention"
            attention_colleges.append(college)
        
        if college_return_rate > best_rate and data['borrowed'] > 0:
            best_rate = college_return_rate
            best_college = college
        
        officers_list = ', '.join(data['accountable_officers'].keys())
        
        ws_summary.cell(row=row_num, column=1, value=f"{college}:")
        row_num += 1
        ws_summary.cell(row=row_num, column=2, value=f"  • Transactions: {data['count']} | Borrowed: {data['borrowed']} | Returned: {data['returned']} | Pending: {data['pending']}")
        row_num += 1
        ws_summary.cell(row=row_num, column=2, value=f"  • Return Rate: {college_return_rate:.1f}% ({rating})")
        row_num += 1
        ws_summary.cell(row=row_num, column=2, value=f"  • Accountable Officer(s): {officers_list}")
        row_num += 1
        ws_summary.cell(row=row_num, column=1, value="")
        row_num += 1
    
    # Key Insights
    ws_summary.cell(row=row_num, column=1, value="KEY INSIGHTS:").font = Font(bold=True, size=12, color='000000')
    row_num += 1
    
    if best_college:
        ws_summary.cell(row=row_num, column=2, value=f"• Best Performing: {best_college} with a {best_rate:.1f}% return rate.")
        row_num += 1
    
    most_active = max(summary_data.items(), key=lambda x: x[1]['count']) if summary_data else (None, None)
    if most_active and most_active[0]:
        ws_summary.cell(row=row_num, column=2, value=f"• Most Active: {most_active[0]} with {most_active[1]['count']} borrowing transaction(s).")
        row_num += 1
    
    if attention_colleges:
        ws_summary.cell(row=row_num, column=2, value=f"• Needs Attention: {', '.join(attention_colleges)} have return rates below 70%.")
        row_num += 1
    
    ws_summary.cell(row=row_num, column=2, value=f"• Overall Return Rate: {overall_return_rate:.1f}% ({total_returned} of {total_borrowed} items).")
    row_num += 1
    ws_summary.cell(row=row_num, column=2, value=f"• Outstanding Items: {total_pending} items still need to be returned.")
    row_num += 2
    
    # Recommendations
    ws_summary.cell(row=row_num, column=1, value="RECOMMENDATIONS:").font = Font(bold=True, size=12, color='000000')
    row_num += 1
    
    if total_pending > 10:
        ws_summary.cell(row=row_num, column=2, value=f"• Follow up on {total_pending} outstanding items across all colleges.")
        row_num += 1
    
    for college in attention_colleges:
        data = summary_data.get(college, {})
        ws_summary.cell(row=row_num, column=2, value=f"• Schedule follow-up with {college} regarding {data.get('pending', 0)} pending item(s).")
        row_num += 1
    
    if overall_return_rate < 80:
        ws_summary.cell(row=row_num, column=2, value="• Consider implementing stricter borrowing policies to improve return rates.")
        row_num += 1
    
    if total_pending <= 10 and not attention_colleges and overall_return_rate >= 80:
        ws_summary.cell(row=row_num, column=2, value="• All colleges are performing well. Continue current monitoring practices.")
        row_num += 1
    
    # Set column widths
    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 50
    ws_summary.column_dimensions['C'].width = 15
    ws_summary.column_dimensions['D'].width = 15
    
    # ========== SHEET 3: Summary Table ==========
    ws_table = wb.create_sheet('Summary Table')
    ws_table.sheet_properties.tabColor = 'FFFFFF'
    
    # Title
    ws_table.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    ws_table.cell(row=1, column=1, value='QUICK REFERENCE SUMMARY BY COLLEGE').font = Font(bold=True, size=14, color='000000')
    ws_table.cell(row=1, column=1).alignment = Alignment(horizontal='center')
    
    # Headers
    headers_bg = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    headers_font = Font(bold=True, color='000000', size=11)
    
    table_headers = ['College / Office', 'Accountable Officer(s)', 'Transactions', 'Borrowed', 'Returned', 'Pending', 'Return Rate']
    for col, header in enumerate(table_headers, start=1):
        cell = ws_table.cell(row=3, column=col, value=header)
        cell.fill = headers_bg
        cell.font = headers_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(bottom=Side(style='thin', color='CCCCCC'))
    
    # Data rows
    row_num = 4
    for college, data in sorted(summary_data.items()):
        return_rate = (data['returned'] / data['borrowed'] * 100) if data['borrowed'] > 0 else 0
        officers_list = ', '.join(data['accountable_officers'].keys())
        
        bg_color = 'FFFFFF' if row_num % 2 == 0 else 'F9F9F9'
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row = Font(color='000000', size=10)
        
        ws_table.cell(row=row_num, column=1, value=college).fill = fill_row
        ws_table.cell(row=row_num, column=1).font = font_row
        
        ws_table.cell(row=row_num, column=2, value=officers_list).fill = fill_row
        ws_table.cell(row=row_num, column=2).font = font_row
        
        ws_table.cell(row=row_num, column=3, value=data['count']).fill = fill_row
        ws_table.cell(row=row_num, column=3).font = font_row
        ws_table.cell(row=row_num, column=3).alignment = Alignment(horizontal='center')
        
        ws_table.cell(row=row_num, column=4, value=data['borrowed']).fill = fill_row
        ws_table.cell(row=row_num, column=4).font = font_row
        ws_table.cell(row=row_num, column=4).alignment = Alignment(horizontal='center')
        
        ws_table.cell(row=row_num, column=5, value=data['returned']).fill = fill_row
        ws_table.cell(row=row_num, column=5).font = font_row
        ws_table.cell(row=row_num, column=5).alignment = Alignment(horizontal='center')
        
        ws_table.cell(row=row_num, column=6, value=data['pending']).fill = fill_row
        ws_table.cell(row=row_num, column=6).font = font_row
        ws_table.cell(row=row_num, column=6).alignment = Alignment(horizontal='center')
        
        rate_cell = ws_table.cell(row=row_num, column=7, value=f'{return_rate:.1f}%')
        rate_cell.fill = fill_row
        rate_cell.font = font_row
        rate_cell.alignment = Alignment(horizontal='center')
        
        if return_rate >= 90:
            rate_cell.font = Font(color='00e5a0', bold=True, size=10)
        elif return_rate >= 70:
            rate_cell.font = Font(color='ffb347', bold=True, size=10)
        else:
            rate_cell.font = Font(color='ff4444', bold=True, size=10)
        
        row_num += 1
    
    # Grand total row
    ws_table.cell(row=row_num, column=1, value='GRAND TOTAL')
    ws_table.cell(row=row_num, column=2, value='')
    ws_table.cell(row=row_num, column=3, value=sum(data['count'] for data in summary_data.values()))
    ws_table.cell(row=row_num, column=4, value=total_borrowed)
    ws_table.cell(row=row_num, column=5, value=total_returned)
    ws_table.cell(row=row_num, column=6, value=total_pending)
    ws_table.cell(row=row_num, column=7, value=f'{overall_return_rate:.1f}%')
    
    for col in range(1, 8):
        cell = ws_table.cell(row=row_num, column=col)
        cell.font = Font(bold=True, color='000000', size=10)
        cell.fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')
    
    # Set column widths
    table_col_widths = [25, 45, 15, 12, 12, 12, 15]
    for col, width in enumerate(table_col_widths, start=1):
        ws_table.column_dimensions[get_column_letter(col)].width = width
    
    # Set column widths for main data sheet
    for col, width in enumerate(col_widths, start=1):
        ws_data.column_dimensions[get_column_letter(col)].width = width
    
    ws_data.freeze_panes = 'A4'
    
    return _xl_response(wb, 'borrow_management')


@login_required
def export_device_monitoring(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied
 
    from inventory.models import Transaction
 
    rows = DeviceMonitor.objects.all().order_by('id')
    
    # Build serial to transaction lookup (same logic as device_monitoring view)
    serial_to_tx = {}
    for tx in Transaction.objects.order_by('-borrowed_at'):
        if not tx.serial_number:
            continue
        for sn in [s.strip() for s in tx.serial_number.split(',') if s.strip()]:
            if sn not in serial_to_tx:
                serial_to_tx[sn] = tx
    
    # First, annotate each row with release_status (same as device_monitoring view)
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
    
    # Collect summary data by college/office and device status
    summary_data = {}
    device_status_summary = {
        'serviceable': 0,
        'non_serviceable': 0,
        'sealed': 0,
        'missing': 0,
        'incomplete': 0,
        'released': 0,
        'returned': 0
    }
    
    for row in rows:
        college = row.office_college or 'Unknown'
        
        if college not in summary_data:
            summary_data[college] = {
                'total_devices': 0,
                'serviceable': 0,
                'non_serviceable': 0,
                'sealed': 0,
                'missing': 0,
                'incomplete': 0,
                'released': 0,
                'returned': 0,
                'devices_with_issues': 0
            }
        
        summary_data[college]['total_devices'] += 1
        
        # Count device statuses
        if row.serviceable:
            summary_data[college]['serviceable'] += 1
            device_status_summary['serviceable'] += 1
        if row.non_serviceable:
            summary_data[college]['non_serviceable'] += 1
            device_status_summary['non_serviceable'] += 1
            summary_data[college]['devices_with_issues'] += 1
        if row.sealed:
            summary_data[college]['sealed'] += 1
            device_status_summary['sealed'] += 1
        if row.missing:
            summary_data[college]['missing'] += 1
            device_status_summary['missing'] += 1
            summary_data[college]['devices_with_issues'] += 1
        if row.incomplete:
            summary_data[college]['incomplete'] += 1
            device_status_summary['incomplete'] += 1
            summary_data[college]['devices_with_issues'] += 1
        
        # Count release/return status (using the computed release_status)
        if hasattr(row, 'release_status'):
            if row.release_status == 'Released':
                summary_data[college]['released'] += 1
                device_status_summary['released'] += 1
            elif row.release_status == 'Returned':
                summary_data[college]['returned'] += 1
                device_status_summary['returned'] += 1
    
    total_devices = len(rows)
    total_issues = device_status_summary['non_serviceable'] + device_status_summary['missing'] + device_status_summary['incomplete']
    health_percentage = ((total_devices - total_issues) / total_devices * 100) if total_devices > 0 else 0
    
    # ========== SHEET 1: Device Details ==========
    wb = Workbook()
    ws_details = wb.active
    ws_details.title = 'Device Details'
    ws_details.sheet_properties.tabColor = '00E5A0'
    
    headers = [
        'Box Number', 'College / Office', 'Accountable Person', 'Borrower Type',
        'Accountable Officer', 'Device', 'Serial Number',
        'Serviceable', 'Non-Serviceable', 'Sealed', 'Missing', 'Incomplete',
        'Release / Return', 'Date Returned', 'Remarks', 'Issue',
    ]
    col_widths = [15, 20, 24, 12, 24, 14, 20, 14, 16, 10, 10, 12, 16, 22, 28, 28]
    
    _xl_title(ws_details, 'Device Monitoring Report', len(headers))
    _xl_header(ws_details, 3, headers)
    
    font_yes = Font(bold=True, color='00E5A0', size=10)
    font_no  = Font(color='6B7080', size=10)
    
    for i, row in enumerate(rows, start=1):
        tx = serial_to_tx.get(row.serial_number.strip()) if row.serial_number else None
 
        if tx:
            release_status = 'Returned' if tx.returned_qty >= tx.quantity_borrowed else 'Released'
            date_ret = format_ph_time(tx.returned_at) if tx.returned_at else '—'
        else:
            release_status = getattr(row, 'release_status', '—')
            date_ret = format_ph_time(row.date_returned) if row.date_returned else '—'
 
        borrower_type_display = (
            'Student'  if row.borrower_type == 'student'  else
            'Employee' if row.borrower_type == 'employee' else '—'
        )
 
        bool_vals = [row.serviceable, row.non_serviceable, row.sealed, row.missing, row.incomplete]
        _xl_row(ws_details, i + 3, [
            row.box_number or '—',
            row.office_college or '—',
            row.accountable_person or '—',
            borrower_type_display,
            row.accountable_officer or '—',
            row.device or 'Tablet',
            row.serial_number or '—',
            '✓' if row.serviceable     else '—',
            '✓' if row.non_serviceable else '—',
            '✓' if row.sealed          else '—',
            '✓' if row.missing         else '—',
            '✓' if row.incomplete      else '—',
            release_status,
            date_ret,
            row.remarks or '—',
            row.issue or '—',
        ], even=(i % 2 == 0))
 
        for col_offset, val in enumerate(bool_vals):
            ws_details.cell(row=i + 3, column=8 + col_offset).font = font_yes if val else font_no
    
    for col, width in enumerate(col_widths, start=1):
        ws_details.column_dimensions[get_column_letter(col)].width = width
    ws_details.freeze_panes = 'A4'
    
    # ========== SHEET 2: Executive Summary (Paragraph Format) ==========
    ws_summary = wb.create_sheet('Executive Summary')
    ws_summary.sheet_properties.tabColor = 'FFB347'
    
    # Title
    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_summary.cell(row=1, column=1, value='DEVICE MONITORING EXECUTIVE SUMMARY')
    title_cell.font = Font(bold=True, size=16, color='00E5A0')
    title_cell.alignment = Alignment(horizontal='center')
    
    # Generation date
    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    date_cell = ws_summary.cell(row=2, column=1, value=f'Report Generated: {format_ph_time(timezone.now())}')
    date_cell.font = Font(size=10, color='6B7080')
    date_cell.alignment = Alignment(horizontal='center')
    
    row_offset = 4
    
    # Overview paragraph
    overview_text = (
        f"OVERVIEW: As of {format_ph_time(timezone.now())}, there are a total of "
        f"{total_devices} devices in the monitoring system across all colleges and offices. "
        f"Out of these, {device_status_summary['serviceable']} devices are serviceable "
        f"({(device_status_summary['serviceable']/total_devices*100):.1f}%). "
        f"Currently, {device_status_summary['released']} devices are released/borrowed, "
        f"and {device_status_summary['returned']} devices have been returned."
    )
    
    ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
    overview_cell = ws_summary.cell(row=row_offset, column=1, value=overview_text)
    overview_cell.alignment = Alignment(wrap_text=True, horizontal='left')
    overview_cell.font = Font(size=11)
    ws_summary.row_dimensions[row_offset].height = 80
    
    row_offset += 2
    
    # Device Status Breakdown paragraph
    ws_summary.cell(row=row_offset, column=1, value='DEVICE STATUS BREAKDOWN:')
    ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='00E5A0')
    row_offset += 1
    
    # Good status devices
    good_status = [
        f"• Serviceable: {device_status_summary['serviceable']} devices ({device_status_summary['serviceable']/total_devices*100:.1f}%)",
        f"• Sealed: {device_status_summary['sealed']} devices ({device_status_summary['sealed']/total_devices*100:.1f}%)"
    ]
    
    for status in good_status:
        ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
        cell = ws_summary.cell(row=row_offset, column=1, value=status)
        cell.alignment = Alignment(wrap_text=True, horizontal='left')
        cell.font = Font(size=11, color='00e5a0')
        row_offset += 1
    
    # Issue status devices
    issue_status = []
    if device_status_summary['non_serviceable'] > 0:
        issue_status.append(f"• Non-Serviceable: {device_status_summary['non_serviceable']} devices need repair")
    if device_status_summary['missing'] > 0:
        issue_status.append(f"• Missing: {device_status_summary['missing']} devices are unaccounted for")
    if device_status_summary['incomplete'] > 0:
        issue_status.append(f"• Incomplete: {device_status_summary['incomplete']} devices have missing parts")
    
    if issue_status:
        ws_summary.cell(row=row_offset, column=1, value='DEVICES NEEDING ATTENTION:')
        ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='FF4444')
        row_offset += 1
        
        for status in issue_status:
            ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
            cell = ws_summary.cell(row=row_offset, column=1, value=status)
            cell.alignment = Alignment(wrap_text=True, horizontal='left')
            cell.font = Font(size=11, color='FF4444')
            row_offset += 1
    
    row_offset += 1
    
    # College/Office breakdown paragraph
    ws_summary.cell(row=row_offset, column=1, value='BREAKDOWN BY COLLEGE/OFFICE:')
    ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='00E5A0')
    row_offset += 1
    
    # Find colleges with most issues
    colleges_with_issues = []
    
    for college, data in sorted(summary_data.items()):
        college_health = ((data['total_devices'] - data['devices_with_issues']) / data['total_devices'] * 100) if data['total_devices'] > 0 else 0
        
        paragraph = (
            f"• {college}: {data['total_devices']} total device(s), "
            f"{data['serviceable']} serviceable, {data['non_serviceable']} non-serviceable, "
            f"{data['missing']} missing, {data['incomplete']} incomplete. "
            f"({college_health:.1f}% healthy). "
            f"Currently {data['released']} device(s) are borrowed, {data['returned']} returned."
        )
        
        ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
        cell = ws_summary.cell(row=row_offset, column=1, value=paragraph)
        cell.alignment = Alignment(wrap_text=True, horizontal='left')
        cell.font = Font(size=11)
        ws_summary.row_dimensions[row_offset].height = 40
        
        if data['devices_with_issues'] > 0:
            colleges_with_issues.append(college)
        
        row_offset += 1
    
    row_offset += 1
    
    # Key Insights
    insights_text = f"KEY INSIGHTS:\n"
    insights_text += f"• Overall Device Health: {health_percentage:.1f}% of devices are in good condition.\n"
    
    if device_status_summary['missing'] > 0:
        insights_text += f"• ALERT: {device_status_summary['missing']} device(s) are marked as MISSING. Immediate investigation recommended.\n"
    
    if device_status_summary['non_serviceable'] > 0:
        insights_text += f"• {device_status_summary['non_serviceable']} device(s) need repair/service.\n"
    
    if colleges_with_issues:
        insights_text += f"• Colleges needing attention: {', '.join(colleges_with_issues)}\n"
    
    total_borrowed = sum(data['released'] for data in summary_data.values())
    if total_borrowed > 0:
        insights_text += f"• {total_borrowed} device(s) are currently borrowed and need to be tracked for return."
    
    ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
    insights_cell = ws_summary.cell(row=row_offset, column=1, value=insights_text)
    insights_cell.alignment = Alignment(wrap_text=True, horizontal='left')
    insights_cell.font = Font(size=11)
    ws_summary.row_dimensions[row_offset].height = 120
    
    row_offset += 2
    
    # Recommendations
    ws_summary.cell(row=row_offset, column=1, value='RECOMMENDATIONS:')
    ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='FFB347')
    row_offset += 1
    
    recommendations = []
    
    if device_status_summary['missing'] > 0:
        recommendations.append(f"• IMMEDIATE ACTION: Conduct a physical inventory check for {device_status_summary['missing']} missing device(s).")
    
    if device_status_summary['non_serviceable'] > 0:
        recommendations.append(f"• Schedule repair/maintenance for {device_status_summary['non_serviceable']} non-serviceable device(s).")
    
    if device_status_summary['incomplete'] > 0:
        recommendations.append(f"• Audit {device_status_summary['incomplete']} incomplete device(s) for missing accessories/parts.")
    
    for college in colleges_with_issues:
        data = summary_data.get(college, {})
        if data['devices_with_issues'] > 0:
            recommendations.append(f"• Follow up with {college} regarding {data['devices_with_issues']} device(s) with issues.")
    
    if not recommendations:
        recommendations.append("• All devices are in good condition. Continue regular monitoring and maintenance.")
        recommendations.append("• Maintain current inventory management practices.")
    
    for rec in recommendations:
        ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
        rec_cell = ws_summary.cell(row=row_offset, column=1, value=rec)
        rec_cell.alignment = Alignment(wrap_text=True, horizontal='left')
        rec_cell.font = Font(size=11)
        row_offset += 1
    
    # Set column widths for summary sheet
    for col in range(1, 5):
        ws_summary.column_dimensions[get_column_letter(col)].width = 35
    
    # ========== SHEET 3: Summary Table ==========
    ws_table = wb.create_sheet('Summary Table')
    ws_table.sheet_properties.tabColor = '448AFF'
    
    # Title
    ws_table.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    table_title = ws_table.cell(row=1, column=1, value='DEVICE MONITORING SUMMARY BY COLLEGE')
    table_title.font = Font(bold=True, size=14, color='00E5A0')
    table_title.alignment = Alignment(horizontal='center')
    
    # Headers
    table_headers = [
        'College / Office', 'Total Devices', 'Serviceable', 'Non-Svc', 
        'Sealed', 'Missing', 'Incomplete', 'Healthy %'
    ]
    for col, header in enumerate(table_headers, start=1):
        cell = ws_table.cell(row=3, column=col, value=header)
        cell.fill = PatternFill(start_color='1E2029', end_color='1E2029', fill_type='solid')
        cell.font = Font(bold=True, color='00E5A0', size=11)
        cell.alignment = Alignment(horizontal='center')
    
    # Write table data
    table_row = 4
    for college, data in sorted(summary_data.items()):
        healthy_percentage = ((data['total_devices'] - data['devices_with_issues']) / data['total_devices'] * 100) if data['total_devices'] > 0 else 0
        
        ws_table.cell(row=table_row, column=1, value=college)
        ws_table.cell(row=table_row, column=2, value=data['total_devices'])
        ws_table.cell(row=table_row, column=3, value=data['serviceable'])
        ws_table.cell(row=table_row, column=4, value=data['non_serviceable'])
        ws_table.cell(row=table_row, column=5, value=data['sealed'])
        ws_table.cell(row=table_row, column=6, value=data['missing'])
        ws_table.cell(row=table_row, column=7, value=data['incomplete'])
        ws_table.cell(row=table_row, column=8, value=f'{healthy_percentage:.1f}%')
        
        # Color code healthy percentage
        health_cell = ws_table.cell(row=table_row, column=8)
        if healthy_percentage >= 90:
            health_cell.font = Font(color='00e5a0', bold=True)
        elif healthy_percentage >= 70:
            health_cell.font = Font(color='ffb347', bold=True)
        else:
            health_cell.font = Font(color='ff4444', bold=True)
        
        # Apply alternating row colors
        if table_row % 2 == 0:
            for col in range(1, 9):
                cell = ws_table.cell(row=table_row, column=col)
                cell.fill = PatternFill(start_color='16181F', end_color='16181F', fill_type='solid')
        
        table_row += 1
    
    # Add grand total row
    ws_table.cell(row=table_row, column=1, value='GRAND TOTAL')
    ws_table.cell(row=table_row, column=2, value=total_devices)
    ws_table.cell(row=table_row, column=3, value=device_status_summary['serviceable'])
    ws_table.cell(row=table_row, column=4, value=device_status_summary['non_serviceable'])
    ws_table.cell(row=table_row, column=5, value=device_status_summary['sealed'])
    ws_table.cell(row=table_row, column=6, value=device_status_summary['missing'])
    ws_table.cell(row=table_row, column=7, value=device_status_summary['incomplete'])
    ws_table.cell(row=table_row, column=8, value=f'{health_percentage:.1f}%')
    
    # Style grand total row
    for col in range(1, 9):
        cell = ws_table.cell(row=table_row, column=col)
        cell.font = Font(bold=True, color='00E5A0')
        cell.fill = PatternFill(start_color='1E2029', end_color='1E2029', fill_type='solid')
    
    # Set column widths for table sheet
    table_col_widths = [30, 15, 15, 12, 12, 12, 12, 15]
    for col, width in enumerate(table_col_widths, start=1):
        ws_table.column_dimensions[get_column_letter(col)].width = width
    
    return _xl_response(wb, 'device_monitoring')@login_required
def export_device_monitoring(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied
 
    from inventory.models import Transaction
 
    rows = DeviceMonitor.objects.all().order_by('id')
    
    # Build serial to transaction lookup (same logic as device_monitoring view)
    serial_to_tx = {}
    for tx in Transaction.objects.order_by('-borrowed_at'):
        if not tx.serial_number:
            continue
        for sn in [s.strip() for s in tx.serial_number.split(',') if s.strip()]:
            if sn not in serial_to_tx:
                serial_to_tx[sn] = tx
    
    # First, annotate each row with release_status (same as device_monitoring view)
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
    
    # Collect summary data by college/office and device status
    summary_data = {}
    device_status_summary = {
        'serviceable': 0,
        'non_serviceable': 0,
        'sealed': 0,
        'missing': 0,
        'incomplete': 0,
        'released': 0,
        'returned': 0
    }
    
    for row in rows:
        college = row.office_college or 'Unknown'
        
        if college not in summary_data:
            summary_data[college] = {
                'total_devices': 0,
                'serviceable': 0,
                'non_serviceable': 0,
                'sealed': 0,
                'missing': 0,
                'incomplete': 0,
                'released': 0,
                'returned': 0,
                'devices_with_issues': 0
            }
        
        summary_data[college]['total_devices'] += 1
        
        # Count device statuses
        if row.serviceable:
            summary_data[college]['serviceable'] += 1
            device_status_summary['serviceable'] += 1
        if row.non_serviceable:
            summary_data[college]['non_serviceable'] += 1
            device_status_summary['non_serviceable'] += 1
            summary_data[college]['devices_with_issues'] += 1
        if row.sealed:
            summary_data[college]['sealed'] += 1
            device_status_summary['sealed'] += 1
        if row.missing:
            summary_data[college]['missing'] += 1
            device_status_summary['missing'] += 1
            summary_data[college]['devices_with_issues'] += 1
        if row.incomplete:
            summary_data[college]['incomplete'] += 1
            device_status_summary['incomplete'] += 1
            summary_data[college]['devices_with_issues'] += 1
        
        # Count release/return status (using the computed release_status)
        if hasattr(row, 'release_status'):
            if row.release_status == 'Released':
                summary_data[college]['released'] += 1
                device_status_summary['released'] += 1
            elif row.release_status == 'Returned':
                summary_data[college]['returned'] += 1
                device_status_summary['returned'] += 1
    
    total_devices = len(rows)
    total_issues = device_status_summary['non_serviceable'] + device_status_summary['missing'] + device_status_summary['incomplete']
    health_percentage = ((total_devices - total_issues) / total_devices * 100) if total_devices > 0 else 0
    
    # ========== SHEET 1: Device Details ==========
    wb = Workbook()
    ws_details = wb.active
    ws_details.title = 'Device Details'
    ws_details.sheet_properties.tabColor = 'FFFFFF'
    
    headers = [
        'Box Number', 'College / Office', 'Accountable Person', 'Borrower Type',
        'Accountable Officer', 'Device', 'Serial Number',
        'Serviceable', 'Non-Serviceable', 'Sealed', 'Missing', 'Incomplete',
        'Release / Return', 'Date Returned', 'Remarks', 'Issue',
    ]
    col_widths = [15, 20, 24, 12, 24, 14, 20, 14, 16, 10, 10, 12, 16, 22, 28, 28]
    
    # Custom title with white background, black text
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
    
    # Header row with white background, black text
    fill_header = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_header = Font(bold=True, color='000000', size=11)
    border = Border(bottom=Side(style='thin', color='CCCCCC'))
    align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for col, heading in enumerate(headers, start=1):
        c = ws_details.cell(row=3, column=col, value=heading)
        c.fill = fill_header
        c.font = font_header
        c.border = border
        c.alignment = align
    ws_details.row_dimensions[3].height = 22
    
    # Data rows with white background, black text
    for i, row in enumerate(rows, start=1):
        tx = serial_to_tx.get(row.serial_number.strip()) if row.serial_number else None
 
        if tx:
            release_status = 'Returned' if tx.returned_qty >= tx.quantity_borrowed else 'Released'
            date_ret = format_ph_time(tx.returned_at) if tx.returned_at else '—'
        else:
            release_status = getattr(row, 'release_status', '—')
            date_ret = format_ph_time(row.date_returned) if row.date_returned else '—'
 
        borrower_type_display = (
            'Student'  if row.borrower_type == 'student'  else
            'Employee' if row.borrower_type == 'employee' else '—'
        )
 
        bool_vals = [row.serviceable, row.non_serviceable, row.sealed, row.missing, row.incomplete]
        
        # Alternate row colors: white and light gray
        bg_color = 'FFFFFF' if i % 2 == 0 else 'F9F9F9'
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row = Font(color='000000', size=10)
        border_row = Border(bottom=Side(style='thin', color='EEEEEE'))
        align_row = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        values = [
            row.box_number or '—',
            row.office_college or '—',
            row.accountable_person or '—',
            borrower_type_display,
            row.accountable_officer or '—',
            row.device or 'Tablet',
            row.serial_number or '—',
            '✓' if row.serviceable     else '—',
            '✓' if row.non_serviceable else '—',
            '✓' if row.sealed          else '—',
            '✓' if row.missing         else '—',
            '✓' if row.incomplete      else '—',
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
        
        # Color the checkmarks green
        for col_offset, val in enumerate(bool_vals):
            if val:
                ws_details.cell(row=i + 3, column=8 + col_offset).font = Font(color='00e5a0', bold=True, size=10)
    
    for col, width in enumerate(col_widths, start=1):
        ws_details.column_dimensions[get_column_letter(col)].width = width
    ws_details.freeze_panes = 'A4'
    
    # ========== SHEET 2: Executive Summary (Paragraph Format) ==========
    ws_summary = wb.create_sheet('Executive Summary')
    ws_summary.sheet_properties.tabColor = 'FFFFFF'
    
    # Title
    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_summary.cell(row=1, column=1, value='DEVICE MONITORING EXECUTIVE SUMMARY')
    title_cell.font = Font(bold=True, size=16, color='000000')
    title_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center')
    
    # Generation date
    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    date_cell = ws_summary.cell(row=2, column=1, value=f'Report Generated: {format_ph_time(timezone.now())}')
    date_cell.font = Font(size=10, color='000000')
    date_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    date_cell.alignment = Alignment(horizontal='center')
    
    row_offset = 4
    
    # Overview paragraph
    overview_text = (
        f"OVERVIEW: As of {format_ph_time(timezone.now())}, there are a total of "
        f"{total_devices} devices in the monitoring system across all colleges and offices. "
        f"Out of these, {device_status_summary['serviceable']} devices are serviceable "
        f"({(device_status_summary['serviceable']/total_devices*100):.1f}%). "
        f"Currently, {device_status_summary['released']} devices are released/borrowed, "
        f"and {device_status_summary['returned']} devices have been returned."
    )
    
    ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
    overview_cell = ws_summary.cell(row=row_offset, column=1, value=overview_text)
    overview_cell.alignment = Alignment(wrap_text=True, horizontal='left')
    overview_cell.font = Font(size=11, color='000000')
    overview_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    ws_summary.row_dimensions[row_offset].height = 80
    
    row_offset += 2
    
    # Device Status Breakdown paragraph
    ws_summary.cell(row=row_offset, column=1, value='DEVICE STATUS BREAKDOWN:')
    ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='000000')
    ws_summary.cell(row=row_offset, column=1).fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    row_offset += 1
    
    # Good status devices
    good_status = [
        f"• Serviceable: {device_status_summary['serviceable']} devices ({device_status_summary['serviceable']/total_devices*100:.1f}%)",
        f"• Sealed: {device_status_summary['sealed']} devices ({device_status_summary['sealed']/total_devices*100:.1f}%)"
    ]
    
    for status in good_status:
        ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
        cell = ws_summary.cell(row=row_offset, column=1, value=status)
        cell.alignment = Alignment(wrap_text=True, horizontal='left')
        cell.font = Font(size=11, color='000000')
        cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        row_offset += 1
    
    # Issue status devices
    issue_status = []
    if device_status_summary['non_serviceable'] > 0:
        issue_status.append(f"• Non-Serviceable: {device_status_summary['non_serviceable']} devices need repair")
    if device_status_summary['missing'] > 0:
        issue_status.append(f"• Missing: {device_status_summary['missing']} devices are unaccounted for")
    if device_status_summary['incomplete'] > 0:
        issue_status.append(f"• Incomplete: {device_status_summary['incomplete']} devices have missing parts")
    
    if issue_status:
        ws_summary.cell(row=row_offset, column=1, value='DEVICES NEEDING ATTENTION:')
        ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='000000')
        ws_summary.cell(row=row_offset, column=1).fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        row_offset += 1
        
        for status in issue_status:
            ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
            cell = ws_summary.cell(row=row_offset, column=1, value=status)
            cell.alignment = Alignment(wrap_text=True, horizontal='left')
            cell.font = Font(size=11, color='000000')
            cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
            row_offset += 1
    
    row_offset += 1
    
    # College/Office breakdown paragraph
    ws_summary.cell(row=row_offset, column=1, value='BREAKDOWN BY COLLEGE/OFFICE:')
    ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='000000')
    ws_summary.cell(row=row_offset, column=1).fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    row_offset += 1
    
    # Find colleges with most issues
    colleges_with_issues = []
    
    for college, data in sorted(summary_data.items()):
        college_health = ((data['total_devices'] - data['devices_with_issues']) / data['total_devices'] * 100) if data['total_devices'] > 0 else 0
        
        paragraph = (
            f"• {college}: {data['total_devices']} total device(s), "
            f"{data['serviceable']} serviceable, {data['non_serviceable']} non-serviceable, "
            f"{data['missing']} missing, {data['incomplete']} incomplete. "
            f"({college_health:.1f}% healthy). "
            f"Currently {data['released']} device(s) are borrowed, {data['returned']} returned."
        )
        
        ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
        cell = ws_summary.cell(row=row_offset, column=1, value=paragraph)
        cell.alignment = Alignment(wrap_text=True, horizontal='left')
        cell.font = Font(size=11, color='000000')
        cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        ws_summary.row_dimensions[row_offset].height = 40
        
        if data['devices_with_issues'] > 0:
            colleges_with_issues.append(college)
        
        row_offset += 1
    
    row_offset += 1
    
    # Key Insights
    insights_text = f"KEY INSIGHTS:\n"
    insights_text += f"• Overall Device Health: {health_percentage:.1f}% of devices are in good condition.\n"
    
    if device_status_summary['missing'] > 0:
        insights_text += f"• ALERT: {device_status_summary['missing']} device(s) are marked as MISSING. Immediate investigation recommended.\n"
    
    if device_status_summary['non_serviceable'] > 0:
        insights_text += f"• {device_status_summary['non_serviceable']} device(s) need repair/service.\n"
    
    if colleges_with_issues:
        insights_text += f"• Colleges needing attention: {', '.join(colleges_with_issues)}\n"
    
    total_borrowed = sum(data['released'] for data in summary_data.values())
    if total_borrowed > 0:
        insights_text += f"• {total_borrowed} device(s) are currently borrowed and need to be tracked for return."
    
    ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
    insights_cell = ws_summary.cell(row=row_offset, column=1, value=insights_text)
    insights_cell.alignment = Alignment(wrap_text=True, horizontal='left')
    insights_cell.font = Font(size=11, color='000000')
    insights_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    ws_summary.row_dimensions[row_offset].height = 120
    
    row_offset += 2
    
    # Recommendations
    ws_summary.cell(row=row_offset, column=1, value='RECOMMENDATIONS:')
    ws_summary.cell(row=row_offset, column=1).font = Font(bold=True, size=12, color='000000')
    ws_summary.cell(row=row_offset, column=1).fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    row_offset += 1
    
    recommendations = []
    
    if device_status_summary['missing'] > 0:
        recommendations.append(f"• IMMEDIATE ACTION: Conduct a physical inventory check for {device_status_summary['missing']} missing device(s).")
    
    if device_status_summary['non_serviceable'] > 0:
        recommendations.append(f"• Schedule repair/maintenance for {device_status_summary['non_serviceable']} non-serviceable device(s).")
    
    if device_status_summary['incomplete'] > 0:
        recommendations.append(f"• Audit {device_status_summary['incomplete']} incomplete device(s) for missing accessories/parts.")
    
    for college in colleges_with_issues:
        data = summary_data.get(college, {})
        if data['devices_with_issues'] > 0:
            recommendations.append(f"• Follow up with {college} regarding {data['devices_with_issues']} device(s) with issues.")
    
    if not recommendations:
        recommendations.append("• All devices are in good condition. Continue regular monitoring and maintenance.")
        recommendations.append("• Maintain current inventory management practices.")
    
    for rec in recommendations:
        ws_summary.merge_cells(start_row=row_offset, start_column=1, end_row=row_offset, end_column=4)
        rec_cell = ws_summary.cell(row=row_offset, column=1, value=rec)
        rec_cell.alignment = Alignment(wrap_text=True, horizontal='left')
        rec_cell.font = Font(size=11, color='000000')
        rec_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        row_offset += 1
    
    # Set column widths for summary sheet
    for col in range(1, 5):
        ws_summary.column_dimensions[get_column_letter(col)].width = 35
    
    # ========== SHEET 3: Summary Table ==========
    ws_table = wb.create_sheet('Summary Table')
    ws_table.sheet_properties.tabColor = 'FFFFFF'
    
    # Title
    ws_table.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    table_title = ws_table.cell(row=1, column=1, value='DEVICE MONITORING SUMMARY BY COLLEGE')
    table_title.font = Font(bold=True, size=14, color='000000')
    table_title.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    table_title.alignment = Alignment(horizontal='center')
    
    # Headers
    fill_header_table = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_header_table = Font(bold=True, color='000000', size=11)
    
    table_headers = [
        'College / Office', 'Total Devices', 'Serviceable', 'Non-Svc', 
        'Sealed', 'Missing', 'Incomplete', 'Healthy %'
    ]
    for col, header in enumerate(table_headers, start=1):
        cell = ws_table.cell(row=3, column=col, value=header)
        cell.fill = fill_header_table
        cell.font = font_header_table
        cell.alignment = Alignment(horizontal='center')
    
    # Write table data
    table_row = 4
    for college, data in sorted(summary_data.items()):
        healthy_percentage = ((data['total_devices'] - data['devices_with_issues']) / data['total_devices'] * 100) if data['total_devices'] > 0 else 0
        
        # Alternate row colors: white and light gray
        bg_color = 'FFFFFF' if table_row % 2 == 0 else 'F9F9F9'
        fill_row = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row = Font(color='000000', size=10)
        
        ws_table.cell(row=table_row, column=1, value=college).fill = fill_row
        ws_table.cell(row=table_row, column=1).font = font_row
        ws_table.cell(row=table_row, column=2, value=data['total_devices']).fill = fill_row
        ws_table.cell(row=table_row, column=2).font = font_row
        ws_table.cell(row=table_row, column=3, value=data['serviceable']).fill = fill_row
        ws_table.cell(row=table_row, column=3).font = font_row
        ws_table.cell(row=table_row, column=4, value=data['non_serviceable']).fill = fill_row
        ws_table.cell(row=table_row, column=4).font = font_row
        ws_table.cell(row=table_row, column=5, value=data['sealed']).fill = fill_row
        ws_table.cell(row=table_row, column=5).font = font_row
        ws_table.cell(row=table_row, column=6, value=data['missing']).fill = fill_row
        ws_table.cell(row=table_row, column=6).font = font_row
        ws_table.cell(row=table_row, column=7, value=data['incomplete']).fill = fill_row
        ws_table.cell(row=table_row, column=7).font = font_row
        health_cell = ws_table.cell(row=table_row, column=8, value=f'{healthy_percentage:.1f}%')
        health_cell.fill = fill_row
        health_cell.font = font_row
        
        # Color code healthy percentage
        if healthy_percentage >= 90:
            health_cell.font = Font(color='00e5a0', bold=True, size=10)
        elif healthy_percentage >= 70:
            health_cell.font = Font(color='ffb347', bold=True, size=10)
        else:
            health_cell.font = Font(color='ff4444', bold=True, size=10)
        
        table_row += 1
    
    # Add grand total row
    ws_table.cell(row=table_row, column=1, value='GRAND TOTAL')
    ws_table.cell(row=table_row, column=2, value=total_devices)
    ws_table.cell(row=table_row, column=3, value=device_status_summary['serviceable'])
    ws_table.cell(row=table_row, column=4, value=device_status_summary['non_serviceable'])
    ws_table.cell(row=table_row, column=5, value=device_status_summary['sealed'])
    ws_table.cell(row=table_row, column=6, value=device_status_summary['missing'])
    ws_table.cell(row=table_row, column=7, value=device_status_summary['incomplete'])
    ws_table.cell(row=table_row, column=8, value=f'{health_percentage:.1f}%')
    
    # Style grand total row
    for col in range(1, 9):
        cell = ws_table.cell(row=table_row, column=col)
        cell.font = Font(bold=True, color='000000', size=10)
        cell.fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
    
    # Set column widths for table sheet
    table_col_widths = [30, 15, 15, 12, 12, 12, 12, 15]
    for col, width in enumerate(table_col_widths, start=1):
        ws_table.column_dimensions[get_column_letter(col)].width = width
    
    return _xl_response(wb, 'device_monitoring')