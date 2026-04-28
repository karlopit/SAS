"""
inventory/context_processors.py

Injects `pending_count` and `graduation_warning_count` into every template context.
Runs only for authenticated staff users to avoid unnecessary DB queries.
"""
from inventory.models import BorrowRequest, Transaction


def graduation_warning_count(request):
    """
    Returns both pending borrow requests count and graduation warning count.
    (Kept the original function name to stay compatible with existing settings.)
    """
    pending_count = 0
    graduation_warning_count = 0

    if request.user.is_authenticated and hasattr(request.user, 'role') and request.user.role == 'staff':
        # Fast count of pending borrow requests
        pending_count = BorrowRequest.objects.filter(status='pending').count()

        # Count active transactions from graduating students (same logic as views)
        graduating_keywords = ['4th', 'fourth', '5th', 'fifth']
        active_trans = Transaction.objects.filter(
            status='borrowed',
            borrow_request__borrower_type='student',
        ).select_related('borrow_request')

        for tx in active_trans:
            br = tx.borrow_request
            if not br:
                continue
            year_level = (br.year_level or br.year_section or '').strip().lower()
            if any(k in year_level for k in graduating_keywords):
                graduation_warning_count += 1

    return {
        'pending_count': pending_count,
        'graduation_warning_count': graduation_warning_count,
    }