"""
inventory/migrations/0021_performance_indexes.py

Adds database indexes on the most-queried fields to dramatically speed up
- borrow_management page (filters by status, borrow_request FK)
- graduation_warnings (filters by status + borrower_type + year_level)
- device_monitoring (order_by box_number)
- dashboard stat counts
- context_processor queries (runs on EVERY page load)
"""
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('inventory', '0020_alter_devicemonitor_is_released'),
    ]

    operations = [
        # Transaction.status — filtered on almost every query
        migrations.AddIndex(
            model_name='transaction',
            index=models.Index(fields=['status'], name='inv_tx_status_idx'),
        ),
        # Transaction.borrowed_at — used for ordering
        migrations.AddIndex(
            model_name='transaction',
            index=models.Index(fields=['-borrowed_at'], name='inv_tx_borrowed_at_idx'),
        ),
        # Composite: status + borrowed_at (covers the most common query pattern)
        migrations.AddIndex(
            model_name='transaction',
            index=models.Index(fields=['status', '-borrowed_at'], name='inv_tx_status_borrowed_idx'),
        ),
        # BorrowRequest.status — pending count runs on every page load
        migrations.AddIndex(
            model_name='borrowrequest',
            index=models.Index(fields=['status'], name='inv_br_status_idx'),
        ),
        # BorrowRequest.borrower_type — used in graduation_warnings join
        migrations.AddIndex(
            model_name='borrowrequest',
            index=models.Index(fields=['borrower_type'], name='inv_br_borrower_type_idx'),
        ),
        # BorrowRequest.created_at — ordering for borrow_requests page
        migrations.AddIndex(
            model_name='borrowrequest',
            index=models.Index(fields=['-created_at'], name='inv_br_created_at_idx'),
        ),
        # DeviceMonitor.is_released — used in dashboard released/returned counts
        migrations.AddIndex(
            model_name='devicemonitor',
            index=models.Index(fields=['is_released'], name='inv_dm_is_released_idx'),
        ),
        # DeviceMonitor.date_returned — used in dashboard + device_monitoring
        migrations.AddIndex(
            model_name='devicemonitor',
            index=models.Index(fields=['date_returned'], name='inv_dm_date_returned_idx'),
        ),
        # DeviceMonitor.office_college — used in bar chart aggregation
        migrations.AddIndex(
            model_name='devicemonitor',
            index=models.Index(fields=['office_college'], name='inv_dm_office_idx'),
        ),
        # DeviceMonitor.box_number — used for ordering (text, so prefix index)
        migrations.AddIndex(
            model_name='devicemonitor',
            index=models.Index(fields=['box_number'], name='inv_dm_box_number_idx'),
        ),
        # TransactionDevice.serial_number — used in return_devices lookup
        migrations.AddIndex(
            model_name='transactiondevice',
            index=models.Index(fields=['serial_number'], name='inv_td_serial_idx'),
        ),
        # TransactionDevice.returned — used in return modal
        migrations.AddIndex(
            model_name='transactiondevice',
            index=models.Index(fields=['returned'], name='inv_td_returned_idx'),
        ),
        # TransactionDevice composite: transaction + returned
        migrations.AddIndex(
            model_name='transactiondevice',
            index=models.Index(fields=['transaction', 'returned'], name='inv_td_tx_returned_idx'),
        ),
    ]