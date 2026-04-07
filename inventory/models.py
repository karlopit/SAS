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

class Transaction(models.Model):
    STATUS_CHOICES = [
        ('borrowed', 'Borrowed'),
        ('returned', 'Returned'),
    ]
    item = models.ForeignKey(Item, on_delete=models.CASCADE, related_name='transactions')
    borrower = models.ForeignKey(CustomUser, on_delete=models.CASCADE, related_name='transactions')
    office_college = models.CharField(max_length=255)
    quantity_borrowed = models.PositiveIntegerField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='borrowed')
    borrowed_at = models.DateTimeField(auto_now_add=True)
    returned_at = models.DateTimeField(null=True, blank=True)
    # Staff-editable condition columns (int only)
    serviceable = models.PositiveIntegerField(default=0)
    unserviceable = models.PositiveIntegerField(default=0)
    sealed = models.PositiveIntegerField(default=0)
    lent_to_students = models.PositiveIntegerField(default=0)
    box_only = models.PositiveIntegerField(default=0)

    def __str__(self):
        return f"{self.borrower.username} borrowed {self.item.name}"