from django.contrib import admin
from django.urls import path, include
from django.contrib.auth.views import LogoutView

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('inventory.urls')),  # Inventory app
    path('users/', include('users.urls')),  # Users app
    path('logout/', LogoutView.as_view(next_page='welcome'), name='logout')
]