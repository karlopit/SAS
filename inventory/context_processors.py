from django.db.models import Q


def graduation_warning_count(request):
    if not request.user.is_authenticated or request.user.role != 'staff':
        return {'graduation_warning_count': 0}

    from inventory.models import BorrowRequest

    graduating_keywords = ['4th', '4', 'fourth', '5th', '5', 'fifth']

    # Get all active borrow requests for students that have active transactions
    active_borrow_requests = BorrowRequest.objects.filter(
        borrower_type='student',
        transaction__status='borrowed',  # Only active transactions
    ).select_related('transaction')

    count = 0
    for br in active_borrow_requests:
        year_level = (br.year_level or br.year_section or '').strip().lower()
        if any(k in year_level for k in graduating_keywords):
            count += 1

    return {'graduation_warning_count': count}