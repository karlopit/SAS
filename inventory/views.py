import random
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.utils import timezone
from .models import Item, Transaction, BorrowRequest
from .forms import ItemForm, StaffBorrowForm, TransactionConditionForm, BorrowRequestForm
from .decorators import no_cache

def welcome(request):
    # If already logged in, go straight to dashboard
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
            # Ensure unique
            while BorrowRequest.objects.filter(transaction_id=tx_id).exists():
                tx_id = str(random.randint(10000, 99999))
            req.transaction_id = tx_id
            req.save()
            borrow_success = req.transaction_id
            generated_tx_id = str(random.randint(10000, 99999))  # fresh one for next use
            borrow_form = BorrowRequestForm()  # reset form
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
    items = Item.objects.all()
    transactions = Transaction.objects.all().order_by('-borrowed_at')[:20]
    pending_count = BorrowRequest.objects.filter(status='pending').count()
    return render(request, 'inventory/index.html', {
        'items': items,
        'transactions': transactions,
        'active_borrows': Transaction.objects.filter(status='borrowed').count(),
        'total_returns': Transaction.objects.filter(status='returned').count(),
        'pending_count': pending_count,
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
def staff_confirm_borrow(request, request_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    borrow_req = get_object_or_404(BorrowRequest, id=request_id, status='pending')

    if request.method == 'POST':
        form = StaffBorrowForm(request.POST)
        if form.is_valid():
            transaction = form.save(commit=False)
            transaction.borrower = request.user
            transaction.borrow_request = borrow_req
            transaction.office_college = borrow_req.office_college
            transaction.status = 'borrowed'
            transaction.item.available_quantity -= transaction.quantity_borrowed
            transaction.item.save()
            transaction.save()
            borrow_req.status = 'accepted'
            borrow_req.save()
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
    return redirect('borrow_requests')

@login_required
def return_item(request, transaction_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    transaction = get_object_or_404(Transaction, id=transaction_id)
    if request.method == 'POST' and transaction.status == 'borrowed':
        transaction.status = 'returned'
        transaction.returned_at = timezone.now()
        transaction.item.available_quantity += transaction.quantity_borrowed
        transaction.item.save()
        transaction.save()
        return redirect('index')
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
    return redirect('index')