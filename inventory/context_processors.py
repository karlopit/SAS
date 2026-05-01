"""
inventory/context_processors.py

Performance fix: this runs on EVERY page load for every authenticated staff user.
The original version did a Python loop over all active transactions — very slow.

Fixes:
1. Uses values() query instead of loading full model objects
2. Caches the result for 30 seconds per user to avoid repeated DB hits
   (3 staff members × every page = 3x the queries)
"""
from django.core.cache import cache
from inventory.models import BorrowRequest, Transaction


def graduation_warning_count(request):
    pending_count = 0
    graduation_warning_count_val = 0

    if not (request.user.is_authenticated and
            hasattr(request.user, 'role') and
            request.user.role == 'staff'):
        return {
            'pending_count': pending_count,
            'graduation_warning_count': graduation_warning_count_val,
        }

    # Cache key is per-user so one user's stale badge doesn't affect others
    cache_key = f'ctx_badges_{request.user.id}'
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    # Fast count with index — no Python loop
    pending_count = BorrowRequest.objects.filter(status='pending').count()

    graduating_keywords = ['4th', 'fourth', '5th', 'fifth']

    # Use values() to avoid loading full objects — only pull the two fields we need
    rows = Transaction.objects.filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).values(
        'borrow_request__year_level',
        'borrow_request__year_section',
    )

    for row in rows:
        yl = (
            row['borrow_request__year_level'] or
            row['borrow_request__year_section'] or ''
        ).strip().lower()
        if any(k in yl for k in graduating_keywords):
            graduation_warning_count_val += 1

    result = {
        'pending_count': pending_count,
        'graduation_warning_count': graduation_warning_count_val,
    }

    # Cache for 30 seconds — stale by at most one page load cycle
    # WebSocket pushes keep the live badges accurate anyway
    cache.set(cache_key, result, 30)

    return result