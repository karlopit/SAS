from django.urls import re_path
from inventory import consumers

websocket_urlpatterns = [
    # Dashboard — stats, pie chart, pending count
    re_path(r'^ws/dashboard/$', consumers.DashboardConsumer.as_asgi()),

    # Borrow Management — transaction table live updates
    re_path(r'^ws/borrow-management/$', consumers.BorrowManagementConsumer.as_asgi()),

    # Borrow Requests — pending request list
    re_path(r'^ws/borrow-requests/$', consumers.BorrowRequestsConsumer.as_asgi()),

    # Device Monitoring — device table
    re_path(r'^ws/device-monitoring/$', consumers.DeviceMonitoringConsumer.as_asgi()),
]