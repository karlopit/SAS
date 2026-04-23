"""
inventory/context_processors.py

Injects `graduation_warning_count` into every template context.
Cached per-user for 60 seconds so the DB is not hit on every single page load.
"""
from django.core.cache import cache


def graduation_warning_count(request):
    if not request.user.is_authenticated:
        return {'graduation_warning_count': 0}

    if not hasattr(request.user, 'role') or request.user.role != 'staff':
        return {'graduation_warning_count': 0}

    cache_key = f'grad_warn_count_{request.user.id}'
    count = cache.get(cache_key)

    if count is None:
        from inventory.models import BorrowRequest
        graduating_keywords = ['4th', 'fourth', '5th', 'fifth']

        active_brs = BorrowRequest.objects.filter(
            borrower_type='student',
            transaction__status='borrowed',
        ).values('year_level', 'year_section')

        count = 0
        for br in active_brs:
            yl = (br['year_level'] or br['year_section'] or '').strip().lower()
            if any(k in yl for k in graduating_keywords):
                count += 1

        cache.set(cache_key, count, 60)

    return {'graduation_warning_count': count}