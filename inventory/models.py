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
    borrower_name  = models.CharField(max_length=255)
    borrower_type  = models.CharField(max_length=20, choices=BORROWER_TYPE_CHOICES, null=True, blank=True)
    office_college = models.CharField(max_length=255, blank=True, null=True)
    item           = models.ForeignKey('Item', on_delete=models.SET_NULL, null=True, blank=True, related_name='borrow_requests')
    quantity       = models.PositiveIntegerField()
    status         = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    created_at     = models.DateTimeField(auto_now_add=True)

    # Student-specific fields
    student_id    = models.CharField(max_length=50, null=True, blank=True)
    year_section  = models.CharField(max_length=100, null=True, blank=True)   # kept for legacy data
    year_level    = models.CharField(max_length=50, null=True, blank=True)    # new: e.g. "4th Year"
    section       = models.CharField(max_length=50, null=True, blank=True)    # new: e.g. "A"
    college       = models.CharField(max_length=200, null=True, blank=True)
    academic_year = models.CharField(max_length=50, null=True, blank=True)

    # Employee-specific fields
    employee_id = models.CharField(max_length=50, null=True, blank=True)
    office      = models.CharField(max_length=200, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.transaction_id:
            self.transaction_id = generate_transaction_id()
        # Keep year_section in sync for backward compatibility
        if self.year_level or self.section:
            parts = [p for p in [self.year_level, self.section] if p]
            self.year_section = ' - '.join(parts)
        super().save(*args, **kwargs)

    def is_graduating(self):
        """Returns True if this student is 4th year or higher."""
        if not self.year_level:
            return False
        yl = self.year_level.strip().lower()
        graduating_keywords = ['4th', '4', 'fourth', '5th', '5', 'fifth']
        return any(k in yl for k in graduating_keywords)

    def __str__(self):
        return f"Request #{self.transaction_id} — {self.borrower_name}"


class Transaction(models.Model):
    STATUS_CHOICES = [
        ('borrowed', 'Borrowed'),
        ('returned', 'Returned'),
    ]
    borrow_request    = models.OneToOneField(BorrowRequest, on_delete=models.SET_NULL, null=True, blank=True, related_name='transaction')
    item              = models.ForeignKey(Item, on_delete=models.CASCADE, related_name='transactions')
    borrower          = models.ForeignKey(CustomUser, on_delete=models.CASCADE, related_name='transactions')
    office_college    = models.CharField(max_length=255)
    quantity_borrowed = models.PositiveIntegerField()
    returned_qty      = models.PositiveIntegerField(default=0)
    serial_number     = models.CharField(max_length=100, blank=True, null=True, help_text="Device serial number (comma-separated)")
    status            = models.CharField(max_length=20, choices=STATUS_CHOICES, default='borrowed')
    borrowed_at       = models.DateTimeField(auto_now_add=True)
    returned_at       = models.DateTimeField(null=True, blank=True)
    serviceable       = models.PositiveIntegerField(default=0)
    unserviceable     = models.PositiveIntegerField(default=0)
    sealed            = models.PositiveIntegerField(default=0)
    lent_to_students  = models.PositiveIntegerField(default=0)
    box_only          = models.PositiveIntegerField(default=0)

    def __str__(self):
        return f"{self.borrower.username} borrowed {self.item.name}"


class TransactionDevice(models.Model):
    transaction   = models.ForeignKey(Transaction, on_delete=models.CASCADE, related_name='devices')
    serial_number = models.CharField(max_length=255)
    box_number    = models.CharField(max_length=100, blank=True)
    returned      = models.BooleanField(default=False)
    returned_at   = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.serial_number} / {self.box_number} — {'returned' if self.returned else 'out'}"


class DeviceMonitor(models.Model):
    box_number          = models.CharField(max_length=100, blank=True)
    office_college      = models.CharField(max_length=255, blank=True)
    accountable_person  = models.CharField(max_length=255, blank=True)
    borrower_type       = models.CharField(max_length=20, choices=[('student', 'Student'), ('employee', 'Employee')], null=True, blank=True)
    accountable_officer = models.CharField(max_length=255, blank=True)
    device              = models.CharField(max_length=255, default='Tablet')
    serial_number       = models.CharField(max_length=255, blank=True)
    serviceable         = models.BooleanField(default=False)
    non_serviceable     = models.BooleanField(default=False)
    sealed              = models.BooleanField(default=False)
    missing             = models.BooleanField(default=False)
    incomplete          = models.BooleanField(default=False)
    remarks             = models.TextField(blank=True)
    issue               = models.TextField(blank=True)
    date_returned       = models.DateTimeField(blank=True, null=True)
    created_at          = models.DateTimeField(auto_now_add=True)
    updated_at          = models.DateTimeField(auto_now=True)
    transaction_id      = models.IntegerField(null=True, blank=True, db_index=True)
    assigned_mr         = models.CharField(max_length=100, blank=True, verbose_name="Assigned M.R.")
    ptr                 = models.CharField(max_length=100, blank=True, verbose_name="PTR")
    is_released         = models.BooleanField(default=False, help_text="True if device is currently borrowed/released")

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.box_number} - {self.device} ({self.serial_number})"