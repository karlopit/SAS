from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.utils import timezone
from .models import Item, Transaction
from .forms import ItemForm, BorrowForm, TransactionConditionForm

@login_required
def index(request):
    items = Item.objects.all()
    transactions = Transaction.objects.all().order_by('-borrowed_at')[:10]
    return render(request, 'inventory/index.html', {
        'items': items,
        'transactions': transactions,
        'active_borrows': Transaction.objects.filter(status='borrowed').count(),
        'total_returns': Transaction.objects.filter(status='returned').count(),
    })

@login_required
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
def borrow_item(request):
    if request.user.role == 'admin':
        return redirect('index')

    available_items = Item.objects.filter(available_quantity__gt=0)

    if request.method == 'POST':
        form = BorrowForm(request.POST)
        if form.is_valid():
            transaction = form.save(commit=False)  # don’t save yet
            transaction.borrower = request.user     # set the borrower
            transaction.save()                      # now save to DB

            # Update item’s available quantity
            transaction.item.available_quantity -= transaction.quantity_borrowed
            transaction.item.save()

            return redirect('index')
    else:
        form = BorrowForm()

    return render(request, 'inventory/borrow_item.html', {
        'form': form,
        'available_items': available_items,
    })

@login_required
def return_item(request, transaction_id):
    transaction = get_object_or_404(Transaction, id=transaction_id, borrower=request.user)
    if request.method == 'POST' and transaction.status == 'borrowed':
        transaction.status = 'returned'
        transaction.returned_at = timezone.now()
        transaction.item.available_quantity += transaction.quantity_borrowed
        transaction.item.save()
        transaction.save()
        return redirect('index')
    return render(request, 'inventory/return_item.html', {'transaction': transaction})

@login_required
def update_condition(request, transaction_id):
    if request.user.role != 'staff':
        raise PermissionDenied
    transaction = get_object_or_404(Transaction, id=transaction_id)
    if request.method == 'POST':
        form = TransactionConditionForm(request.POST, instance=transaction)
        if form.is_valid():
            form.save()
    return redirect('index')