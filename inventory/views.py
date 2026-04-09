import json
import random
from django.db import models as db_models
from django.db.models import Sum
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.http import JsonResponse
from django.utils import timezone
from django.views.decorators.http import require_POST
from .models import Item, Transaction, BorrowRequest, DeviceMonitor
from .forms import ItemForm, StaffBorrowForm, TransactionConditionForm, BorrowRequestForm
from .decorators import no_cache


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

    from django.db.models import Sum, F, ExpressionWrapper, IntegerField
    still_out = Transaction.objects.filter(status='borrowed').annotate(
        still_borrowed=ExpressionWrapper(
            F('quantity_borrowed') - F('returned_qty'),
            output_field=IntegerField()
        )
    ).aggregate(total=Sum('still_borrowed'))['total'] or 0
    borrowed_qty = max(0, still_out)

    returned_qty = Transaction.objects.aggregate(
        total=Sum('returned_qty')
    )['total'] or 0

    monitors       = DeviceMonitor.objects.all()
    offices        = list(monitors.values_list('office_college', flat=True).distinct())
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
        'returned_qty':   returned_qty,
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
    items        = Item.objects.all()
    transactions = Transaction.objects.all().order_by('-borrowed_at')[:20]
    pending_count = BorrowRequest.objects.filter(status='pending').count()
    return render(request, 'inventory/borrow_management.html', {
        'items':        items,
        'transactions': transactions,
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
        'rows':         rows,
        'pending_count': pending_count,
    })


@login_required
@require_POST
def device_monitoring_save(request):
    if request.user.role != 'staff':
        raise PermissionDenied

    ids              = request.POST.getlist('row_id')
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

        is_svc  = get(serviceables)      == 'on'
        is_ns   = get(non_serviceables)  == 'on'
        is_seal = get(sealeds)           == 'on'
        is_miss = get(missings)          == 'on'
        is_inc  = get(incompletes)       == 'on'

        if row_id == 'new':
            DeviceMonitor.objects.create(
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
        else:
            try:
                obj = DeviceMonitor.objects.get(pk=int(row_id))
                obj.office_college     = get(offices)
                obj.accountable_person = get(accountables)
                obj.device             = get(devices) or 'Tablet'
                obj.serial_number      = get(serials)
                obj.serviceable        = is_svc
                obj.non_serviceable    = is_ns
                obj.sealed             = is_seal
                obj.missing            = is_miss
                obj.incomplete         = is_inc
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
            transaction.borrower        = request.user
            transaction.borrow_request  = borrow_req
            transaction.office_college  = borrow_req.office_college
            transaction.status          = 'borrowed'
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
        qty_to_restore = transaction.returned_qty  # use what staff entered, not full qty
        if qty_to_restore > 0:
            transaction.item.available_quantity += qty_to_restore
            transaction.item.save()
        # Only fully close the transaction when everything is back
        if transaction.returned_qty >= transaction.quantity_borrowed:
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
    transaction = get_object_or_404(Transaction, id=transaction_id)
    if request.method == 'POST':
        form = TransactionConditionForm(request.POST, instance=transaction)
        if form.is_valid():
            form.save()
    return redirect('borrow_management')