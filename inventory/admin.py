from django.contrib import admin
from .models import Item, Transaction

@admin.register(Item)
class ItemAdmin(admin.ModelAdmin):
    list_display = ['name', 'quantity', 'available_quantity', 'created_at']
    search_fields = ['name']

@admin.register(Transaction)
class TransactionAdmin(admin.ModelAdmin):
    list_display = ['borrower', 'item', 'quantity_borrowed', 'status', 'borrowed_at', 'returned_at']
    list_filter = ['status']