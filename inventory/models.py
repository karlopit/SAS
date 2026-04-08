import random
from django.db import models
from users.models import CustomUser

class Item(models.Model):
    name = models.CharField(max_length=255)
    description = models.TextField(blank=True)
    serial = models.TextField(blank=True)
    quantity = models.PositiveIntegerField(default=0)
    available_quantity = models.PositiveIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

def generate_transaction_id():
    """Generate a unique 5-digit transaction ID."""
    while True:
        tx_id = str(random.randint(10000, 99999))
        if not BorrowRequest.objects.filter(transaction_id=tx_id).exists():
            return tx_id

class BorrowRequest(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('accepted', 'Accepted'),
        ('declined', 'Declined'),
    ]
    transaction_id = models.CharField(max_length=5, unique=True)
    borrower_name = models.CharField(max_length=255)
    office_college = models.CharField(max_length=255)
    item = models.ForeignKey(Item, on_delete=models.SET_NULL, null=True, blank=True, related_name='borrow_requests')
    quantity = models.PositiveIntegerField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    created_at = models.DateTimeField(auto_now_add=True)

    def save(self, *args, **kwargs):
        # Auto-assign unique transaction_id on first save
        if not self.transaction_id:
            self.transaction_id = generate_transaction_id()
        super().save(*args, **kwargs)

    def __str__(self):
        return f"Request #{self.transaction_id} — {self.borrower_name}"

class Transaction(models.Model):
    STATUS_CHOICES = [
        ('borrowed', 'Borrowed'),
        ('returned', 'Returned'),
    ]
    borrow_request = models.OneToOneField(BorrowRequest, on_delete=models.SET_NULL, null=True, blank=True, related_name='transaction')
    item = models.ForeignKey(Item, on_delete=models.CASCADE, related_name='transactions')
    borrower = models.ForeignKey(CustomUser, on_delete=models.CASCADE, related_name='transactions')
    office_college = models.CharField(max_length=255)
    quantity_borrowed = models.PositiveIntegerField()
    returned_qty = models.PositiveIntegerField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='borrowed')
    borrowed_at = models.DateTimeField(auto_now_add=True)
    returned_at = models.DateTimeField(null=True, blank=True)
    serviceable = models.PositiveIntegerField(default=0)
    unserviceable = models.PositiveIntegerField(default=0)
    sealed = models.PositiveIntegerField(default=0)
    lent_to_students = models.PositiveIntegerField(default=0)
    box_only = models.PositiveIntegerField(default=0)

    def __str__(self):
        return f"{self.borrower.username} borrowed {self.item.name}"