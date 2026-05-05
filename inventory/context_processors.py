# inventory/context_processors.py — replace the entire file

from django.core.cache import cache
from inventory.models import BorrowRequest, Transaction

# Shared across all users — same data, no reason for per-user keys
BADGES_CACHE_KEY = 'nav_badges_shared'
BADGES_CACHE_TTL = 300  # 5 minutes; WS broadcasts will invalidate this


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

    cached = cache.get(BADGES_CACHE_KEY)
    if cached is not None:
        return cached

    # Simple count — no annotations, no Case/When, no Lower/Trim
    pending_count = BorrowRequest.objects.filter(status='pending').count()

    # Simplified graduation check — filter in DB with raw LIKE, no annotations
    from django.db.models import Q
    grad_q = (
        Q(borrow_request__year_level__icontains='4th') |
        Q(borrow_request__year_level__icontains='fourth') |
        Q(borrow_request__year_level__icontains='5th') |
        Q(borrow_request__year_level__icontains='fifth') |
        Q(borrow_request__year_section__icontains='4th') |
        Q(borrow_request__year_section__icontains='fourth') |
        Q(borrow_request__year_section__icontains='5th') |
        Q(borrow_request__year_section__icontains='fifth')
    )
    graduation_warning_count_val = Transaction.objects.filter(
        status='borrowed',
        borrow_request__borrower_type='student',
    ).filter(grad_q).count()

    result = {
        'pending_count': pending_count,
        'graduation_warning_count': graduation_warning_count_val,
    }

    cache.set(BADGES_CACHE_KEY, result, BADGES_CACHE_TTL)
    return result


def invalidate_nav_badges():
    """Call this after any mutation that affects pending count or grad warnings."""
    cache.delete(BADGES_CACHE_KEY)