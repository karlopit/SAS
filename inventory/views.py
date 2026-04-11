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
from .models import Item
from django.contrib import messages

def welcome(request):
    if request.user.is_authenticated:
        return redirect('index')

    borrow_form = BorrowRequestForm()
    borrow_success = None
    generated_tx_id = str(random.randint(10000, 99999))

    if request.method == 'POST' and request.POST.get('action') == 'borrow_request':
        borrow_form = BorrowRequestForm(request.POST)
        if borrow_form.is_valid():
            req = borrow_form.save(commit=False)
            tx_id = request.POST.get('transaction_id', str(random.randint(10000, 99999)))
            while BorrowRequest.objects.filter(transaction_id=tx_id).exists():
                tx_id = str(random.randint(10000, 99999))
            req.transaction_id = tx_id
            req.save()
            borrow_success = req.transaction_id
            generated_tx_id = str(random.randint(10000, 99999))
            borrow_form = BorrowRequestForm()
            return redirect('welcome')

    return render(request, 'inventory/welcome.html', {
        'borrow_form': borrow_form,
        'borrow_success': borrow_success,
        'generated_tx_id': generated_tx_id,
        'available_items': Item.objects.filter(available_quantity__gt=0),
    })


@login_required
@no_cache
def index(request):
    pending_count  = BorrowRequest.objects.filter(status='pending').count()
    items          = Item.objects.all()
    active_borrows = Transaction.objects.filter(status='borrowed').count()
    total_returns  = Transaction.objects.filter(status='returned').count()

    available_qty = sum(i.available_quantity for i in items)

    agg = Transaction.objects.annotate(
        still_out=ExpressionWrapper(
            F('quantity_borrowed') - F('returned_qty'),
            output_field=IntegerField()
        )
    ).aggregate(total=Sum('still_out'))
    borrowed_qty = max(0, agg['total'] or 0)

    monitors = DeviceMonitor.objects.all()

    offices = sorted(set(
        monitors.values_list('office_college', flat=True)
    ))

    dm_serviceable = [monitors.filter(office_college=o, serviceable=True).count()     for o in offices]
    dm_non_service = [monitors.filter(office_college=o, non_serviceable=True).count() for o in offices]
    dm_sealed      = [monitors.filter(office_college=o, sealed=True).count()          for o in offices]
    dm_missing     = [monitors.filter(office_college=o, missing=True).count()         for o in offices]
    dm_incomplete  = [monitors.filter(office_college=o, incomplete=True).count()      for o in offices]

    return render(request, 'inventory/index.html', {
        'items':          items,
        'active_borrows': active_borrows,
        'total_returns':  total_returns,
        'pending_count':  pending_count,
        'available_qty':  available_qty,
        'borrowed_qty':   borrowed_qty,
        'dm_offices':     json.dumps(offices),
        'dm_serviceable': json.dumps(dm_serviceable),
        'dm_non_service': json.dumps(dm_non_service),
        'dm_sealed':      json.dumps(dm_sealed),
        'dm_missing':     json.dumps(dm_missing),
        'dm_incomplete':  json.dumps(dm_incomplete),
    })


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
        return redirect('index')
    return render(request, 'inventory/add_item.html', {'form': form})

@login_required
def edit_item(request, item_id):
    """Edit an item's available quantity"""
    # Only admin can edit items
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
                    messages.success(request, f'Successfully updated {item.name} quantity to {item.available_quantity}')
                else:
                    messages.error(request, 'Quantity cannot be negative.')
            except ValueError:
                messages.error(request, 'Invalid quantity value. Please enter a valid number.')
        else:
            messages.error(request, 'No quantity provided.')
        
        return redirect('index')  # Redirect back to dashboard
    
    # If not POST, redirect to dashboard
    return redirect('index')


@login_required
@no_cache
def borrow_requests(request):
    if request.user.role != 'staff':
        raise PermissionDenied
    pending = BorrowRequest.objects.filter(status='pending').order_by('-created_at')
    return render(request, 'inventory/borrow_requests.html', {'pending': pending})


@login_required
@no_cache
def borrow_management(request):
    if request.user.role != 'staff':
        raise PermissionDenied
    items         = Item.objects.all()
    transactions  = Transaction.objects.all().order_by('-borrowed_at')[:20]
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
    display_ids      = request.POST.getlist('display_id')
    offices          = request.POST.getlist('office_college')
    accountables     = request.POST.getlist('accountable_person')
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

        is_svc  = get(serviceables)     == 'on'
        is_ns   = get(non_serviceables) == 'on'
        is_seal = get(sealeds)          == 'on'
        is_miss = get(missings)         == 'on'
        is_inc  = get(incompletes)      == 'on'

        fields = dict(
            display_id         = get(display_ids),
            office_college     = get(offices),
            accountable_person = get(accountables),
            device             = get(devices) or 'Tablet',
            serial_number      = get(serials),
            serviceable        = is_svc,
            non_serviceable    = is_ns,
            sealed             = is_seal,
            missing            = is_miss,
            incomplete         = is_inc,
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

    return redirect('device_monitoring')


@login_required
@require_POST
def device_monitoring_delete(request, row_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    obj = get_object_or_404(DeviceMonitor, pk=row_id)
    obj.delete()
    return redirect('device_monitoring')


@login_required
@no_cache
def staff_confirm_borrow(request, request_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    borrow_req = get_object_or_404(BorrowRequest, id=request_id, status='pending')

    if request.method == 'POST':
        form = StaffBorrowForm(request.POST)
        if form.is_valid():
            transaction = form.save(commit=False)
            transaction.borrower       = request.user
            transaction.borrow_request = borrow_req
            transaction.office_college = borrow_req.office_college
            transaction.status         = 'borrowed'
            transaction.item.available_quantity -= transaction.quantity_borrowed
            transaction.item.save()
            transaction.save()
            borrow_req.status = 'accepted'
            borrow_req.save()
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
    return redirect('borrow_requests')


@login_required
def return_item(request, transaction_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    transaction = get_object_or_404(Transaction, id=transaction_id)
    if request.method == 'POST' and transaction.status != 'returned':
        transaction.status      = 'returned'
        transaction.returned_at = timezone.now()
        transaction.save()
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

    old_returned = tx.returned_qty
    delta = new_returned - old_returned

    if delta != 0:
        tx.item.available_quantity = max(0, tx.item.available_quantity + delta)
        tx.item.save()

    tx.returned_qty = new_returned
    tx.returned_at  = timezone.now() if new_returned > 0 else None
    tx.status       = 'returned' if new_returned >= tx.quantity_borrowed else 'borrowed'
    tx.save()

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
        'pie': {
            'available': available_qty,
            'borrowed':  borrowed_qty,
        }
    })


# ── Shared Excel helper utilities ─────────────────────────────────────────────

def _xl_title(ws, text, col_count):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
    c = ws.cell(row=1, column=1)
    c.value     = text
    c.font      = Font(bold=True, size=14, color='00E5A0')
    c.fill      = PatternFill(start_color='0E0F13', end_color='0E0F13', fill_type='solid')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=col_count)
    s = ws.cell(row=2, column=1)
    s.value     = f'Generated: {timezone.now().strftime("%B %d, %Y  %H:%M")}'
    s.font      = Font(size=9, color='6B7080')
    s.fill      = PatternFill(start_color='0E0F13', end_color='0E0F13', fill_type='solid')
    s.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 16


def _xl_header(ws, row_num, headers):
    fill   = PatternFill(start_color='1E2029', end_color='1E2029', fill_type='solid')
    font   = Font(bold=True, color='00E5A0', size=11)
    border = Border(bottom=Side(style='thin', color='2A2D3A'))
    align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col, heading in enumerate(headers, start=1):
        c            = ws.cell(row=row_num, column=col, value=heading)
        c.fill       = fill
        c.font       = font
        c.border     = border
        c.alignment  = align
    ws.row_dimensions[row_num].height = 22


def _xl_row(ws, row_num, values, even=False):
    bg     = '1A1C24' if even else '16181F'
    fill   = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    font   = Font(color='E8EAF0', size=10)
    border = Border(bottom=Side(style='thin', color='2A2D3A'))
    align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col, val in enumerate(values, start=1):
        c           = ws.cell(row=row_num, column=col, value=val)
        c.fill      = fill
        c.font      = font
        c.border    = border
        c.alignment = align


def _xl_response(wb, filename_prefix):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f'{filename_prefix}_{timezone.now().strftime("%Y%m%d_%H%M")}.xlsx'
    resp = HttpResponse(
        buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp


# ── Export: Borrow Management ─────────────────────────────────────────────────

@login_required
def export_borrow_management(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied

    transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).all().order_by('-borrowed_at')

    headers = [
        'Tx ID', 'Borrower Name', 'College / Office',
        'Item', 'Serial Number', 'Qty Borrowed',
        'Returned Qty', 'Borrowed On', 'Returned On',
    ]
    col_widths = [12, 24, 20, 22, 20, 14, 14, 18, 22]

    wb = Workbook()
    ws = wb.active
    ws.title = 'Borrow Management'
    ws.sheet_properties.tabColor = '00E5A0'

    _xl_title(ws, 'Borrow Management Report', len(headers))
    _xl_header(ws, 3, headers)

    for i, tx in enumerate(transactions, start=1):
        tx_id       = f'#{tx.borrow_request.transaction_id}' if tx.borrow_request else '—'
        borrower    = tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower.username
        borrowed_on = tx.borrowed_at.strftime('%b %d, %Y')
        returned_on = tx.returned_at.strftime('%b %d, %Y  %H:%M') if tx.returned_at else '—'

        _xl_row(ws, i + 3, [
            tx_id,
            borrower,
            tx.office_college or '—',
            tx.item.name,
            tx.item.serial or '—',
            tx.quantity_borrowed,
            tx.returned_qty,
            borrowed_on,
            returned_on,
        ], even=(i % 2 == 0))

    for col, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = 'A4'
    return _xl_response(wb, 'borrow_management')


# ── Export: Device Monitoring ─────────────────────────────────────────────────

@login_required
def export_device_monitoring(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied

    rows = DeviceMonitor.objects.all().order_by('id')

    headers = [
        'ID', 'College / Office', 'Accountable Person',
        'Device', 'Serial Number',
        'Serviceable', 'Non-Serviceable', 'Sealed', 'Missing', 'Incomplete',
    ]
    col_widths = [10, 20, 24, 14, 20, 14, 16, 10, 10, 12]

    wb = Workbook()
    ws = wb.active
    ws.title = 'Device Monitoring'
    ws.sheet_properties.tabColor = '00E5A0'

    _xl_title(ws, 'Device Monitoring Report', len(headers))
    _xl_header(ws, 3, headers)

    font_yes = Font(bold=True, color='00E5A0', size=10)
    font_no  = Font(color='6B7080', size=10)

    for i, row in enumerate(rows, start=1):
        bool_vals = [
            row.serviceable,
            row.non_serviceable,
            row.sealed,
            row.missing,
            row.incomplete,
        ]
        _xl_row(ws, i + 3, [
            row.display_id or str(row.id),
            row.office_college or '—',
            row.accountable_person or '—',
            row.device or 'Tablet',
            row.serial_number or '—',
            '✓' if row.serviceable     else '—',
            '✓' if row.non_serviceable else '—',
            '✓' if row.sealed          else '—',
            '✓' if row.missing         else '—',
            '✓' if row.incomplete      else '—',
        ], even=(i % 2 == 0))

        # Apply green/muted colour to the boolean columns (6–10)
        for col_offset, val in enumerate(bool_vals):
            ws.cell(row=i + 3, column=6 + col_offset).font = font_yes if val else font_no

    for col, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = 'A4'
    return _xl_response(wb, 'device_monitoring')