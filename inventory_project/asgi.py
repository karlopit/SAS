"""
ASGI config — Django Channels entry point.
WebSocket connections are routed here; HTTP falls back to Django.
"""
import os
import django
from django.core.asgi import get_asgi_application

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'inventory_project.settings')
django.setup()

from channels.routing import ProtocolTypeRouter, URLRouter
from channels.auth import AuthMiddlewareStack
from channels.security.websocket import AllowedHostsOriginValidator
import inventory.routing

application = ProtocolTypeRouter({
    # Standard HTTP requests
    'http': get_asgi_application(),

    # WebSocket connections — wrapped with auth + allowed-hosts check
    'websocket': AllowedHostsOriginValidator(
        AuthMiddlewareStack(
            URLRouter(inventory.routing.websocket_urlpatterns)
        )
    ),
})