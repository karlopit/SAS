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
    BORROWER_TYPE_CHOICES = [
        ('student', 'Student'),
        ('employee', 'Employee'),
    ]
    
    transaction_id = models.CharField(max_length=5, unique=True)
    borrower_name = models.CharField(max_length=255)
    borrower_type = models.CharField(max_length=20, choices=BORROWER_TYPE_CHOICES, null=True, blank=True)
    office_college = models.CharField(max_length=255, blank=True, null=True)
    item = models.ForeignKey(Item, on_delete=models.SET_NULL, null=True, blank=True, related_name='borrow_requests')
    quantity = models.PositiveIntegerField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    created_at = models.DateTimeField(auto_now_add=True)
    
    # Student-specific fields
    student_id = models.CharField(max_length=50, null=True, blank=True)
    year_section = models.CharField(max_length=100, null=True, blank=True)
    college = models.CharField(max_length=200, null=True, blank=True)
    academic_year = models.CharField(max_length=50, null=True, blank=True)
    
    # Employee-specific fields
    employee_id = models.CharField(max_length=50, null=True, blank=True)
    office = models.CharField(max_length=200, null=True, blank=True)

    def save(self, *args, **kwargs):
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
    serial_number = models.CharField(max_length=100, blank=True, null=True, help_text="Device serial number")
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


class DeviceMonitor(models.Model):
    display_id          = models.CharField(max_length=100, blank=True)
    office_college      = models.CharField(max_length=255, blank=True)
    accountable_person  = models.CharField(max_length=255, blank=True)
    device              = models.CharField(max_length=255, default='Tablet')
    serial_number       = models.CharField(max_length=255, blank=True)
    serviceable         = models.BooleanField(default=False)
    non_serviceable     = models.BooleanField(default=False)
    sealed              = models.BooleanField(default=False)
    missing             = models.BooleanField(default=False)
    incomplete          = models.BooleanField(default=False)
    created_at          = models.DateTimeField(auto_now_add=True)
    updated_at          = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.office_college} — {self.device} ({self.serial_number})"