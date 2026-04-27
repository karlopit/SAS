"""
inventory/context_processors.py

Injects `graduation_warning_count` into every template context.
Counts active transactions of 4th/5th year students (same logic as the graduation_warnings view).
No caching – ensures the badge is always correct on every request.
"""
from django.core.cache import cache


def graduation_warning_count(request):
    if not request.user.is_authenticated:
        return {'graduation_warning_count': 0}

    if not hasattr(request.user, 'role') or request.user.role != 'staff':
        return {'graduation_warning_count': 0}

    # No caching – simply compute fresh count every time.
    # The query is light (two filtered joins) and acceptable for a small office.
    from inventory.models import Transaction, BorrowRequest

    graduating_keywords = ['4th', 'fourth', '5th', 'fifth']

    # Get all active transactions (status='borrowed') that belong to a student borrow request
    active_transactions = Transaction.objects.filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).select_related('borrow_request')

    count = 0
    for tx in active_transactions:
        br = tx.borrow_request
        if not br:
            continue
        year_level = (br.year_level or br.year_section or '').strip().lower()
        if any(k in year_level for k in graduating_keywords):
            count += 1

    return {'graduation_warning_count': count}