from functools import wraps

def no_cache(view_func):
    """
    Private, non-shared cache semantics without ``no-store``.

    ``no-store`` blocks the browser back/forward cache (bfcache), so every
    in-app navigation pays a full network round-trip. Staff pages still must
    not be cached on shared proxies; ``private, max-age=0`` revalidates while
    allowing bfcache when the browser can restore the tab instantly.
    """
    def wrapper(request, *args, **kwargs):
        response = view_func(request, *args, **kwargs)
        if response is None:
            return response
        response['Cache-Control'] = 'private, max-age=0'
        return response
    return wrapper