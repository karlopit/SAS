from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('inventory.urls')),  # Inventory app
    path('users/', include('users.urls')),  # Users app
]