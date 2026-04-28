import io
import json
import random
import pytz
import openpyxl
import traceback
import re
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
from .forms import ItemForm, StaffBorrowForm, TransactionConditionForm, BorrowRequestForm
from .decorators import no_cache
from django.contrib import messages
from django.views.decorators.csrf import ensure_csrf_cookie, csrf_exempt
from django.views.decorators.http import require_http_methods
from datetime import datetime, date as _date


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
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)

    tx = get_object_or_404(Transaction, id=transaction_id)
    devices = list(tx.devices.all())

    if not devices and tx.serial_number:
        serials = [s.strip() for s in tx.serial_number.split(',') if s.strip()]
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

@ensure_csrf_cookie
def borrow_item_public(request):
    if request.method == 'POST':
        form = BorrowRequestForm(request.POST)
        if form.is_valid():
            br = form.save(commit=False)
            br.save()
            messages.success(request, 'Your request has been submitted. Staff will review it soon.')
            return redirect('borrow_item_public')
    else:
        form = BorrowRequestForm()
    
    return render(request, 'inventory/borrow_item.html', {
        'form': form,
        'available_items': Item.objects.filter(available_quantity__gt=0),
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
 
    # ── Build lookup maps in ONE query each instead of per-row queries ────────
    # Map serial_number → active TransactionDevice (unreturned)
    active_tds = TransactionDevice.objects.filter(
        returned=False
    ).select_related('transaction', 'transaction__borrow_request')
 
    active_serial_map = {}
    for td in active_tds:
        sn = td.serial_number
        if sn and sn not in active_serial_map:
            active_serial_map[sn] = td
 
    # Annotate each row using the pre-built map — zero extra DB queries
    for row in rows:
        sn = (row.serial_number or '').strip()
 
        if row.date_returned:
            row.release_status        = 'Returned'
            row.date_returned_display = format_ph_time(row.date_returned)
        elif sn and sn in active_serial_map:
            td = active_serial_map[sn]
            tx = td.transaction
            tx_borrower = tx.borrow_request.borrower_name if tx.borrow_request else tx.borrower_id
 
            if str(tx_borrower) == row.accountable_person and tx.office_college == row.office_college:
                row.release_status = 'Released'
            else:
                row.release_status = '—'
            row.date_returned_display = '—'
        else:
            row.release_status        = '—'
            row.date_returned_display = '—'
 
    pending_count = BorrowRequest.objects.filter(status='pending').count()
 
    return render(request, 'inventory/device_monitoring.html', {
        'rows':          rows,
        'pending_count': pending_count,
    })
 
 
# ─────────────────────────────────────────────────────────────────────────────
# REPLACE: graduation_warnings view  (was Python-looping all transactions)
# ─────────────────────────────────────────────────────────────────────────────
 
@login_required
@no_cache
def graduation_warnings(request):
    if request.user.role != 'staff':
        raise PermissionDenied
 
    graduating_keywords = ['4th', 'fourth', '5th', 'fifth']
 
    # Filter as much as possible in the DB query
    active_transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).order_by('-borrowed_at')
 
    # Prefetch all devices in one query instead of per-transaction
    from django.db.models import Prefetch
    active_transactions = active_transactions.prefetch_related(
        Prefetch('devices', queryset=TransactionDevice.objects.all())
    )
 
    warnings = []
    for tx in active_transactions:
        br = tx.borrow_request
        if not br:
            continue
        year_level = (br.year_level or '').strip().lower()
        if not year_level:
            year_level = (br.year_section or '').strip().lower()
        if not any(k in year_level for k in graduating_keywords):
            continue
 
        qty_outstanding = tx.quantity_borrowed - tx.returned_qty
 
        # devices already prefetched — no extra query here
        all_devices = tx.devices.all()
        if all_devices:
            serials_display = ', '.join(d.serial_number for d in all_devices)
        else:
            serials_display = tx.serial_number or '—'
 
        warnings.append({
            'borrower_name':   br.borrower_name,
            'year_level':      br.year_level or br.year_section or '—',
            'section':         br.section or '—',
            'college':         br.college or br.office_college or '—',
            'academic_year':   br.academic_year or '—',
            'student_id':      br.student_id or '—',
            'item_name':       tx.item.name,
            'qty_outstanding': qty_outstanding,
            'serial_number':   serials_display,
            'borrowed_at':     format_ph_time(tx.borrowed_at),
            'officer':         (tx.borrower.get_full_name() or '').strip() or tx.borrower.username,
            'tx_id':           br.transaction_id,
        })
 
    pending_count = BorrowRequest.objects.filter(status='pending').count()
 
    return render(request, 'inventory/graduation_warnings.html', {
        'warnings':      warnings,
        'warning_count': len(warnings),
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
    assigned_mr_list     = request.POST.getlist('assigned_mr')
    accountables         = request.POST.getlist('accountable_person')
    borrower_types       = request.POST.getlist('borrower_type')
    accountable_officers = request.POST.getlist('accountable_officer')
    devices              = request.POST.getlist('device')
    serials              = request.POST.getlist('serial_number')
    serviceables         = request.POST.getlist('serviceable')
    non_serviceables     = request.POST.getlist('non_serviceable')
    sealeds              = request.POST.getlist('sealed')
    missings             = request.POST.getlist('missing')
    ptr_list             = request.POST.getlist('ptr')
    incompletes          = request.POST.getlist('incomplete')
    remarks_list         = request.POST.getlist('remarks')
    issue_list           = request.POST.getlist('issue')

    for i, row_id in enumerate(ids):
        def get(lst, idx=i):
            return lst[idx] if idx < len(lst) else ''

        fields = dict(
            box_number          = get(box_numbers),
            office_college      = get(offices),
            assigned_mr         = get(assigned_mr_list),
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
            ptr                 = get(ptr_list),
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
#  Helper: normalize a header string for fuzzy matching
# ─────────────────────────────────────────────────────────────────────────────
def _normalize_header(h):
    """
    'Assigned M.R. #' → 'assigned mr'
    'P.T.R.'          → 'ptr'
    'College / Office'→ 'college / office'   (spaces around / kept for the slash-variants)
    Strip, lowercase, remove ALL dots and hashes, collapse whitespace.
    """
    h = str(h or '').strip().lower()
    h = re.sub(r'\.', '', h)          # remove all full-stops
    h = re.sub(r'#', '', h)           # remove hash signs
    h = re.sub(r'\s+', ' ', h).strip()
    return h
 
 
# ─────────────────────────────────────────────────────────────────────────────
#  Helper: parse an Excel cell value into a PH-timezone-aware datetime
# ─────────────────────────────────────────────────────────────────────────────
def _parse_excel_date(raw):
    """
    Convert an openpyxl cell value to a timezone-aware datetime (Asia/Manila).
    Returns None for blank / unparseable values.
    """
    if raw is None or str(raw).strip() in ('', '—', '-', 'N/A', 'None'):
        return None
    if isinstance(raw, datetime):
        return raw if raw.tzinfo else PH_TZ.localize(raw)
    if isinstance(raw, _date):
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
 
 
# ─────────────────────────────────────────────────────────────────────────────
#  The import view
# ─────────────────────────────────────────────────────────────────────────────
@login_required
@require_http_methods(["POST"])
def device_monitoring_import(request):
    if request.user.role != 'staff':
        return JsonResponse({'error': 'Forbidden'}, status=403)
 
    try:
        excel_file = request.FILES.get('excel_file')
        if not excel_file:
            return JsonResponse({'error': 'No file provided'}, status=400)
        if not excel_file.name.endswith(('.xlsx', '.xls')):
            return JsonResponse({'error': 'Invalid file format. Use .xlsx or .xls'}, status=400)
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        ws = wb.active
    except Exception as e:
        return JsonResponse({'error': f'Excel read error: {str(e)}'}, status=400)
 
    # ── Locate the header row ─────────────────────────────────────────────────
    # We look for a row that contains a cell whose NORMALIZED text is one of
    # the serial-number aliases.
    SERIAL_NORM = {'serial no', 's/n', 'serial number', 'serial'}
    header_row_num = None
    header_row_raw = []   # original strings (for the debug response)
    header_row_norm = []  # normalized strings (for mapping)
 
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        raw_cells  = [str(cell or '') for cell in row]
        norm_cells = [_normalize_header(cell) for cell in raw_cells]
        if any(nc in SERIAL_NORM for nc in norm_cells):
            header_row_num  = row_idx
            header_row_raw  = raw_cells
            header_row_norm = norm_cells
            break
 
    if header_row_num is None:
        return JsonResponse({
            'error': 'Could not find a header row with a Serial Number column in the first 10 rows.'
        }, status=400)
 
    # ── HEADER MAP (keys are NORMALIZED strings) ──────────────────────────────
    #
    # After _normalize_header():
    #   "Assigned M.R."   → "assigned mr"
    #   "Assigned M.R. #" → "assigned mr"
    #   "Assigned MR"     → "assigned mr"
    #   "PTR"             → "ptr"
    #   "PTR #"           → "ptr"
    #   "P.T.R."          → "ptr"
    #
    HEADER_MAP = {
        # box_number
        'box no':     'box_number',
        'box number': 'box_number',
        'box':        'box_number',
        # serial_number
        'serial no':     'serial_number',
        'serial number': 'serial_number',
        's/n':           'serial_number',
        'serial':        'serial_number',
        # office_college
        'college/office':  'office_college',
        'college / office':'office_college',
        'office':          'office_college',
        'college':         'office_college',
        # accountable_person
        'name of student':  'accountable_person',
        'name':             'accountable_person',
        'student name':     'accountable_person',
        'accountable person': 'accountable_person',
        # borrower_type
        'borrower type': 'borrower_type',
        'type':          'borrower_type',
        # accountable_officer
        'accountable officer': 'accountable_officer',
        'officer':             'accountable_officer',
        # ── assigned_mr ───────────────────────────────────────────────────────
        # After normalization all these collapse to 'assigned mr' or 'mr':
        'assigned mr':   'assigned_mr',   # "Assigned M.R.", "Assigned M.R. #", "Assigned MR", etc.
        'assigned m r':  'assigned_mr',   # in case of odd spacing
        'mr':            'assigned_mr',   # standalone "M.R." or "MR"
        # ── ptr ───────────────────────────────────────────────────────────────
        # After normalization: "PTR", "PTR #", "P.T.R." all → 'ptr'
        'ptr':           'ptr',
        'property tag':  'ptr',
        # device
        'device': 'device',
        # date_returned
        'date returned': 'date_returned',
        'return date':   'date_returned',
        'returned date': 'date_returned',
        'returned on':   'date_returned',
        # release/return status
        'release / return':  'release_status_import',
        'release/return':    'release_status_import',
        'release status':    'release_status_import',
        'released/returned': 'release_status_import',
        # remarks / issue
        'remarks': 'remarks',
        'issue':   'issue',
        # 'status' intentionally NOT mapped — we use date_returned +
        # release_status_import instead of the checkbox-style status column.
    }
 
    # ── Build column-index → field-name map ───────────────────────────────────
    col_map = {}
    for idx, norm in enumerate(header_row_norm):
        field = HEADER_MAP.get(norm)
        if field:
            col_map[idx] = field
 
    if 'serial_number' not in col_map.values():
        return JsonResponse({
            'error': (
                'Serial column found but could not be mapped. '
                f'Raw headers detected: {header_row_raw}'
            )
        }, status=400)
 
    has_date_col = 'date_returned' in col_map.values()
 
    created = 0
    updated = 0
    errors  = []
 
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=header_row_num + 1, values_only=True),
        start=header_row_num + 1,
    ):
        try:
            if not row or all(cell is None for cell in row):
                continue
 
            data     = {}   # string values
            raw_data = {}   # raw values (needed for date parsing)
            for col_idx, field_name in col_map.items():
                if col_idx < len(row):
                    raw = row[col_idx]
                    raw_data[field_name] = raw
                    data[field_name] = str(raw).strip() if raw is not None else ''
 
            serial_number = data.get('serial_number', '').strip()
            if not serial_number:
                continue
 
            # ── borrower_type ─────────────────────────────────────────────────
            bt_raw = data.get('borrower_type', '').lower()
            borrower_type = (
                'student'  if bt_raw == 'student'  else
                'employee' if bt_raw == 'employee' else ''
            )
 
            # ── date_returned ─────────────────────────────────────────────────
            date_returned = _parse_excel_date(raw_data.get('date_returned'))
 
            # ── release_status_import → informs date_returned ─────────────────
            release_text = data.get('release_status_import', '').strip().lower()
            if release_text == 'returned':
                if date_returned is None:
                    date_returned = get_ph_time()
            elif release_text in ('released', '—', '-', ''):
                if not has_date_col:
                    date_returned = None
 
            # ── upsert ────────────────────────────────────────────────────────
            _obj, created_flag = DeviceMonitor.objects.update_or_create(
                serial_number=serial_number,
                defaults={
                    'box_number':          data.get('box_number', ''),
                    'office_college':      data.get('office_college', ''),
                    'accountable_person':  data.get('accountable_person', ''),
                    'borrower_type':       borrower_type,
                    'accountable_officer': data.get('accountable_officer', ''),
                    'assigned_mr':         data.get('assigned_mr', ''),
                    'device':              data.get('device', '') or 'Tablet',
                    'ptr':                 data.get('ptr', ''),
                    'remarks':             data.get('remarks', ''),
                    'issue':              data.get('issue', ''),
                    'date_returned':       date_returned,
                    # Checkboxes always reset to unchecked on import
                    'serviceable':         False,
                    'non_serviceable':     False,
                    'sealed':              False,
                    'missing':             False,
                    'incomplete':          False,
                }
            )
            if created_flag:
                created += 1
            else:
                updated += 1
 
        except Exception as e:
            errors.append(f'Row {row_idx}: {str(e)}')
 
    b = _broadcasts()
    b.broadcast_device_monitoring()
 
    return JsonResponse({
        'ok':               True,
        'created':          created,
        'updated':          updated,
        'errors':           errors,
        # Included so you can verify exact header strings detected from your Excel:
        'headers_detected': header_row_raw,
        'headers_mapped':   {str(k): v for k, v in col_map.items()},
    })


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
            assigned_mr = request.POST.get('assigned_mr', '').strip()

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
                    assigned_mr=assigned_mr,
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
    except (json.JSONDecodeError, AttributeError):
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    device_ids = body.get('device_ids', [])
    serials    = body.get('serials', [])

    now_ph = get_ph_time()
    returned_serials = []

    # ── Mark TransactionDevice rows as returned ───────────────────────────────
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
                td.returned    = True
                td.returned_at = now_ph
                td.save()
                returned_serials.append(sn)

    # ── Mirror the return into DeviceMonitor (sets date_returned) ────────────
    if returned_serials:
        if tx.borrow_request:
            borrower_name = tx.borrow_request.borrower_name
            office        = tx.borrow_request.office_college
        else:
            borrower_name = tx.borrower.get_full_name() or tx.borrower.username
            office        = tx.office_college

        DeviceMonitor.objects.filter(
            serial_number__in=returned_serials,
            accountable_person=borrower_name,
            office_college=office,
            date_returned__isnull=True,
        ).update(date_returned=now_ph)

    # ── Recalculate returned count ────────────────────────────────────────────
    if tx.devices.exists():
        returned_count = tx.devices.filter(returned=True).count()
    else:
        # Legacy path: no TransactionDevice rows
        returned_count = tx.returned_qty + len(returned_serials)

    returned_count = min(returned_count, tx.quantity_borrowed)

    # ── Update Transaction ────────────────────────────────────────────────────
    delta = returned_count - tx.returned_qty
    if delta > 0:
        tx.item.available_quantity = tx.item.available_quantity + delta
        tx.item.save()

    tx.returned_qty = returned_count
    tx.returned_at  = now_ph if returned_count > 0 else tx.returned_at
    tx.status       = 'returned' if returned_count >= tx.quantity_borrowed else 'borrowed'
    tx.save()

    # ── Broadcast live updates to ALL connected pages ─────────────────────────
    b = _broadcasts()
    b.broadcast_borrow_management()
    b.broadcast_dashboard()
    b.broadcast_device_monitoring()   # ← device monitoring page updates live

    return JsonResponse({
        'ok':            True,
        'returned_qty':  tx.returned_qty,
        'status':        tx.status,
        'fully_returned': tx.returned_qty >= tx.quantity_borrowed,
        'returned_at':   format_ph_time(tx.returned_at),
    })


# ─────────────────────────────────────────────────────────────────────────────
#  Graduation Warnings
# ─────────────────────────────────────────────────────────────────────────────

@login_required
@no_cache
def graduation_warnings(request):
    """
    Shows staff a list of active borrowers who are 4th year (or higher),
    meaning they are near graduation and their tablets should be recalled.
    Displays all serial numbers of devices borrowed in the transaction,
    regardless of whether they have been returned or not.
    """
    if request.user.role != 'staff':
        raise PermissionDenied

    graduating_keywords = ['4th', '4', 'fourth', '5th', '5', 'fifth']

    active_transactions = Transaction.objects.select_related(
        'item', 'borrower', 'borrow_request'
    ).filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).order_by('-borrowed_at')

    warnings = []
    for tx in active_transactions:
        br = tx.borrow_request
        if not br:
            continue
        year_level = (br.year_level or '').strip().lower()
        if not year_level:
            year_level = (br.year_section or '').strip().lower()
        if not any(k in year_level for k in graduating_keywords):
            continue

        qty_outstanding = tx.quantity_borrowed - tx.returned_qty

        # --- Get ALL serial numbers (including returned) ---
        all_devices = tx.devices.all()
        if all_devices.exists():
            # Use TransactionDevice records (all, regardless of returned flag)
            all_serials = [d.serial_number for d in all_devices]
            serials_display = ', '.join(all_serials)
        else:
            # Fallback for legacy transactions: use the comma-separated field
            serials_display = tx.serial_number or '—'

        warnings.append({
            'borrower_name':   br.borrower_name,
            'year_level':      br.year_level or br.year_section or '—',
            'section':         br.section or '—',
            'college':         br.college or br.office_college or '—',
            'academic_year':   br.academic_year or '—',
            'student_id':      br.student_id or '—',
            'item_name':       tx.item.name,
            'qty_outstanding': qty_outstanding,
            'serial_number':   serials_display,
            'borrowed_at':     format_ph_time(tx.borrowed_at),
            'officer':         (tx.borrower.get_full_name() or '').strip() or tx.borrower.username,
            'tx_id':           br.transaction_id,
        })

    pending_count = BorrowRequest.objects.filter(status='pending').count()

    return render(request, 'inventory/graduation_warnings.html', {
        'warnings':      warnings,
        'warning_count': len(warnings),
        'pending_count': pending_count,
    })


# ─────────────────────────────────────────────────────────────────────────────
#  Excel exports — helpers
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


# ─────────────────────────────────────────────────────────────────────────────
#  Export: Borrow Management
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def export_borrow_management(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied

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

    # ── Sheet 1: Transaction Details ──────────────────────────────────────────
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

    # Collect summary data
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

    # Totals
    total_borrowed      = sum(d['borrowed'] for d in summary_data.values())
    total_returned      = sum(d['returned'] for d in summary_data.values())
    total_pending       = sum(d['pending']  for d in summary_data.values())
    overall_return_rate = (total_returned / total_borrowed * 100) if total_borrowed > 0 else 0

    # ── Sheet 2: Summary Report ───────────────────────────────────────────────
    ws_summary = wb.create_sheet('Summary Report')
    ws_summary.sheet_properties.tabColor = 'FFFFFF'

    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_summary.cell(row=1, column=1, value='BORROW MANAGEMENT SUMMARY REPORT')
    title_cell.font = Font(bold=True, size=16, color='000000')
    title_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center')

    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    date_cell = ws_summary.cell(row=2, column=1, value=f'Report Generated: {format_ph_time(timezone.now())}')
    date_cell.font = Font(size=10, color='000000')
    date_cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    date_cell.alignment = Alignment(horizontal='center')

    row_num = 4

    ws_summary.cell(row=row_num, column=1, value='OVERVIEW:').font = Font(bold=True, size=12, color='000000')
    row_num += 1

    overview_text = (
        f"As of {format_ph_time(timezone.now())}, there have been a total of "
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

    # ── Sheet 3: Summary Table ────────────────────────────────────────────────
    ws_table = wb.create_sheet('Summary Table')
    ws_table.sheet_properties.tabColor = 'FFFFFF'

    ws_table.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    tbl_title = ws_table.cell(row=1, column=1, value='QUICK REFERENCE SUMMARY BY COLLEGE')
    tbl_title.font = Font(bold=True, size=14, color='000000')
    tbl_title.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    tbl_title.alignment = Alignment(horizontal='center')

    tbl_headers     = ['College / Office', 'Accountable Officer(s)', 'Transactions', 'Borrowed', 'Returned', 'Pending', 'Return Rate']
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

    # Grand total
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

    return _xl_response(wb, 'borrow_management')


# ─────────────────────────────────────────────────────────────────────────────
#  Export: Device Monitoring
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def export_device_monitoring(request):
    if request.user.role not in ('staff', 'admin'):
        raise PermissionDenied

    rows = DeviceMonitor.objects.all().order_by('id')

    # Annotate release_status on each row
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

    # Collect summary data
    summary_data = {}
    device_status_summary = {
        'serviceable': 0, 'non_serviceable': 0, 'sealed': 0,
        'missing': 0, 'incomplete': 0, 'released': 0, 'returned': 0,
    }

    for row in rows:
        college = row.office_college or 'Unknown'
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

    total_devices    = len(rows)
    total_issues     = (device_status_summary['non_serviceable']
                        + device_status_summary['missing']
                        + device_status_summary['incomplete'])
    health_percentage = ((total_devices - total_issues) / total_devices * 100) if total_devices > 0 else 0

    wb = Workbook()

    # ── Sheet 1: Device Details ───────────────────────────────────────────────
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

    ws_details.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws_details.cell(row=1, column=1, value='Device Monitoring Report')
    c.font      = Font(bold=True, size=14, color='000000')
    c.fill      = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws_details.row_dimensions[1].height = 30

    ws_details.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    s = ws_details.cell(row=2, column=1)
    ph_now  = get_ph_time()
    s.value = f'Generated: {ph_now.strftime("%B %d, %Y %I:%M %p")}'
    s.font  = Font(size=9, color='000000')
    s.fill  = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    s.alignment = Alignment(horizontal='center', vertical='center')
    ws_details.row_dimensions[2].height = 16

    fill_hdr = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_hdr = Font(bold=True, color='000000', size=11)
    bdr      = Border(bottom=Side(style='thin', color='CCCCCC'))
    aln      = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for col, heading in enumerate(headers, start=1):
        cell = ws_details.cell(row=3, column=col, value=heading)
        cell.fill = fill_hdr; cell.font = font_hdr
        cell.border = bdr;    cell.alignment = aln
    ws_details.row_dimensions[3].height = 22

    for i, row in enumerate(rows, start=1):
        borrower_type_display = (
            'Student'  if row.borrower_type == 'student'  else
            'Employee' if row.borrower_type == 'employee' else '—'
        )
        release_status = getattr(row, 'release_status', '—')
        date_ret = format_ph_time(row.date_returned) if row.date_returned else '—'

        bg_color   = 'FFFFFF' if i % 2 == 0 else 'F9F9F9'
        fill_row   = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_row   = Font(color='000000', size=10)
        border_row = Border(bottom=Side(style='thin', color='EEEEEE'))
        align_row  = Alignment(horizontal='center', vertical='center', wrap_text=True)

        bool_vals = [row.serviceable, row.non_serviceable, row.sealed, row.missing, row.incomplete]
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
            cell.fill = fill_row; cell.font = font_row
            cell.border = border_row; cell.alignment = align_row

        for col_offset, val in enumerate(bool_vals):
            if val:
                ws_details.cell(row=i + 3, column=8 + col_offset).font = Font(color='00e5a0', bold=True, size=10)

    for col, width in enumerate(col_widths, start=1):
        ws_details.column_dimensions[get_column_letter(col)].width = width
    ws_details.freeze_panes = 'A4'

    # ── Sheet 2: Summary Report ───────────────────────────────────────────────
    ws_summary = wb.create_sheet('Summary Report')
    ws_summary.sheet_properties.tabColor = 'FFFFFF'

    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws_summary.cell(row=1, column=1, value='DEVICE MONITORING SUMMARY REPORT')
    title_cell.font      = Font(bold=True, size=16, color='000000')
    title_cell.fill      = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center')

    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    date_cell = ws_summary.cell(row=2, column=1, value=f'Report Generated: {format_ph_time(timezone.now())}')
    date_cell.font      = Font(size=10, color='000000')
    date_cell.fill      = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    date_cell.alignment = Alignment(horizontal='center')

    row_offset = 4

    def _sum_write(ws, row, text, bold=False, color='000000', height=None):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=1, value=text)
        cell.alignment = Alignment(wrap_text=True, horizontal='left')
        cell.font = Font(size=11 if not bold else 12, color=color, bold=bold)
        cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        if height:
            ws.row_dimensions[row].height = height

    # Overview
    svc_pct = (device_status_summary['serviceable'] / total_devices * 100) if total_devices > 0 else 0
    overview_text = (
        f"OVERVIEW: As of {format_ph_time(timezone.now())}, there are a total of "
        f"{total_devices} devices in the monitoring system across all colleges and offices. "
        f"Out of these, {device_status_summary['serviceable']} devices are serviceable "
        f"({svc_pct:.1f}%). "
        f"Currently, {device_status_summary['released']} devices are released/borrowed, "
        f"and {device_status_summary['returned']} devices have been returned."
    )
    _sum_write(ws_summary, row_offset, overview_text, height=80)
    row_offset += 2

    _sum_write(ws_summary, row_offset, 'DEVICE STATUS BREAKDOWN:', bold=True)
    row_offset += 1

    for line in [
        f'• Serviceable: {device_status_summary["serviceable"]} devices ({svc_pct:.1f}%)',
        f'• Sealed: {device_status_summary["sealed"]} devices ({(device_status_summary["sealed"]/total_devices*100) if total_devices else 0:.1f}%)',
    ]:
        _sum_write(ws_summary, row_offset, line)
        row_offset += 1

    issue_lines = []
    if device_status_summary['non_serviceable'] > 0:
        issue_lines.append(f'• Non-Serviceable: {device_status_summary["non_serviceable"]} devices need repair')
    if device_status_summary['missing'] > 0:
        issue_lines.append(f'• Missing: {device_status_summary["missing"]} devices are unaccounted for')
    if device_status_summary['incomplete'] > 0:
        issue_lines.append(f'• Incomplete: {device_status_summary["incomplete"]} devices have missing parts')

    if issue_lines:
        _sum_write(ws_summary, row_offset, 'DEVICES NEEDING ATTENTION:', bold=True)
        row_offset += 1
        for line in issue_lines:
            _sum_write(ws_summary, row_offset, line)
            row_offset += 1

    row_offset += 1
    _sum_write(ws_summary, row_offset, 'BREAKDOWN BY COLLEGE/OFFICE:', bold=True)
    row_offset += 1

    colleges_with_issues = []
    for college, data in sorted(summary_data.items()):
        college_health = ((data['total_devices'] - data['devices_with_issues']) / data['total_devices'] * 100) if data['total_devices'] > 0 else 0
        officers_list  = ', '.join(set(
            (tx.borrower.get_full_name() or '').strip() or tx.borrower.username
            for tx in Transaction.objects.filter(office_college=college).select_related('borrower')
        )) or '—'

        paragraph = (
            f"• {college}: {data['total_devices']} total device(s), "
            f"{data['serviceable']} serviceable, {data['non_serviceable']} non-serviceable, "
            f"{data['missing']} missing, {data['incomplete']} incomplete. "
            f"({college_health:.1f}% healthy). "
            f"Currently {data['released']} device(s) are borrowed, {data['returned']} returned. "
            f"Accountable Officer(s): {officers_list}."
        )
        _sum_write(ws_summary, row_offset, paragraph, height=50)
        if data['devices_with_issues'] > 0:
            colleges_with_issues.append(college)
        row_offset += 1

    row_offset += 1
    insights_lines = [
        f'KEY INSIGHTS:',
        f'• Overall Device Health: {health_percentage:.1f}% of devices are in good condition.',
    ]
    if device_status_summary['missing'] > 0:
        insights_lines.append(f'• ALERT: {device_status_summary["missing"]} device(s) are marked as MISSING. Immediate investigation recommended.')
    if device_status_summary['non_serviceable'] > 0:
        insights_lines.append(f'• {device_status_summary["non_serviceable"]} device(s) need repair/service.')
    if colleges_with_issues:
        insights_lines.append(f'• Colleges needing attention: {", ".join(colleges_with_issues)}')
    total_borrowed_dm = sum(d['released'] for d in summary_data.values())
    if total_borrowed_dm > 0:
        insights_lines.append(f'• {total_borrowed_dm} device(s) are currently borrowed and need to be tracked for return.')

    _sum_write(ws_summary, row_offset, '\n'.join(insights_lines), height=120)
    row_offset += 2

    _sum_write(ws_summary, row_offset, 'RECOMMENDATIONS:', bold=True)
    row_offset += 1

    recs = []
    if device_status_summary['missing'] > 0:
        recs.append(f'• IMMEDIATE ACTION: Conduct a physical inventory check for {device_status_summary["missing"]} missing device(s).')
    if device_status_summary['non_serviceable'] > 0:
        recs.append(f'• Schedule repair/maintenance for {device_status_summary["non_serviceable"]} non-serviceable device(s).')
    if device_status_summary['incomplete'] > 0:
        recs.append(f'• Audit {device_status_summary["incomplete"]} incomplete device(s) for missing accessories/parts.')
    for college in colleges_with_issues:
        recs.append(f'• Follow up with {college} regarding {summary_data[college]["devices_with_issues"]} device(s) with issues.')
    if not recs:
        recs.append('• All devices are in good condition. Continue regular monitoring and maintenance.')
        recs.append('• Maintain current inventory management practices.')

    for line in recs:
        _sum_write(ws_summary, row_offset, line)
        row_offset += 1

    for col in range(1, 5):
        ws_summary.column_dimensions[get_column_letter(col)].width = 35

    # ── Sheet 3: Summary Table ────────────────────────────────────────────────
    ws_table = wb.create_sheet('Summary Table')
    ws_table.sheet_properties.tabColor = 'FFFFFF'

    ws_table.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    tbl_title = ws_table.cell(row=1, column=1, value='DEVICE MONITORING SUMMARY BY COLLEGE')
    tbl_title.font      = Font(bold=True, size=14, color='000000')
    tbl_title.fill      = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    tbl_title.alignment = Alignment(horizontal='center')

    fill_tbl_hdr = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
    font_tbl_hdr = Font(bold=True, color='000000', size=11)

    tbl_headers = ['College / Office', 'Total Devices', 'Serviceable', 'Non-Svc',
                   'Sealed', 'Missing', 'Incomplete', 'Healthy %']
    for col, hdr in enumerate(tbl_headers, start=1):
        cell = ws_table.cell(row=3, column=col, value=hdr)
        cell.fill = fill_tbl_hdr; cell.font = font_tbl_hdr
        cell.alignment = Alignment(horizontal='center')

    tbl_row = 4
    for college, data in sorted(summary_data.items()):
        healthy_pct = ((data['total_devices'] - data['devices_with_issues']) / data['total_devices'] * 100) if data['total_devices'] > 0 else 0

        bg_color = 'FFFFFF' if tbl_row % 2 == 0 else 'F9F9F9'
        fill_r   = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        font_r   = Font(color='000000', size=10)

        row_vals = [college, data['total_devices'], data['serviceable'], data['non_serviceable'],
                    data['sealed'], data['missing'], data['incomplete'], f'{healthy_pct:.1f}%']
        for col, val in enumerate(row_vals, start=1):
            cell = ws_table.cell(row=tbl_row, column=col, value=val)
            cell.fill = fill_r; cell.font = font_r
            cell.alignment = Alignment(horizontal='center')

        health_cell = ws_table.cell(row=tbl_row, column=8)
        if healthy_pct >= 90:
            health_cell.font = Font(color='00e5a0', bold=True, size=10)
        elif healthy_pct >= 70:
            health_cell.font = Font(color='ffb347', bold=True, size=10)
        else:
            health_cell.font = Font(color='ff4444', bold=True, size=10)

        tbl_row += 1

    grand_vals = ['GRAND TOTAL', total_devices,
                  device_status_summary['serviceable'], device_status_summary['non_serviceable'],
                  device_status_summary['sealed'], device_status_summary['missing'],
                  device_status_summary['incomplete'], f'{health_percentage:.1f}%']
    for col, val in enumerate(grand_vals, start=1):
        cell = ws_table.cell(row=tbl_row, column=col, value=val)
        cell.font = Font(bold=True, color='000000', size=10)
        cell.fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')

    tbl_col_widths = [30, 15, 15, 12, 12, 12, 12, 15]
    for col, width in enumerate(tbl_col_widths, start=1):
        ws_table.column_dimensions[get_column_letter(col)].width = width

    return _xl_response(wb, 'device_monitoring')