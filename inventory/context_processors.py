"""
inventory/context_processors.py

Runs on every authenticated staff page load. Keep it to cheap aggregate queries
and cache — slow work here delays every navigation (especially on cloud DBs).
"""
from django.core.cache import cache
from django.db.models import Case, CharField, F, Q, When
from django.db.models.functions import Lower, Trim
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

    cache_key = f'ctx_badges_{request.user.id}'
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    pending_count = BorrowRequest.objects.filter(status='pending').count()

    # Mirror graduation_warnings view: trimmed year_level if set, else year_section
    kw_q = Q()
    for kw in ('4th', 'fourth', '5th', 'fifth'):
        kw_q |= Q(_eff__icontains=kw)

    graduation_warning_count_val = (
        Transaction.objects.filter(
            status='borrowed',
            borrow_request__borrower_type='student',
        )
        .annotate(
            _ylt=Trim('borrow_request__year_level'),
            _yst=Trim('borrow_request__year_section'),
        )
        .annotate(
            _eff=Lower(
                Case(
                    When(
                        Q(borrow_request__year_level__isnull=True)
                        | Q(_ylt__isnull=True)
                        | Q(_ylt=''),
                        then=F('_yst'),
                    ),
                    default=F('_ylt'),
                    output_field=CharField(),
                )
            )
        )
        .filter(kw_q)
        .count()
    )

    result = {
        'pending_count': pending_count,
        'graduation_warning_count': graduation_warning_count_val,
    }

    # WebSocket updates refresh badges; longer TTL cuts DB wake-ups on Render/Neon
    cache.set(cache_key, result, 120)

    return result