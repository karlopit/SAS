from django.urls import path
from . import views

urlpatterns = [
    # ── Public ───────────────────────────────────────────────────────────────
    path('', views.welcome, name='welcome'),

    # ── Dashboard ─────────────────────────────────────────────────────────────
    path('dashboard/', views.index, name='index'),

    # ── AJAX poll fallback (used when WebSocket is unavailable) ───────────────
    path('ajax/dashboard/',          views.ajax_dashboard_data,          name='ajax_dashboard'),
    path('ajax/borrow-management/',  views.ajax_borrow_management_data,  name='ajax_borrow_management'),
    path('ajax/borrow-requests/',    views.ajax_borrow_requests_data,    name='ajax_borrow_requests'),
    path('ajax/device-monitoring/',  views.ajax_device_monitoring_data,  name='ajax_device_monitoring'),

    # ── Item management ───────────────────────────────────────────────────────
    path('add/', views.add_item, name='add_item'),
    path('item/<int:item_id>/edit/', views.edit_item, name='edit_item'),

    # ── Borrow flow ───────────────────────────────────────────────────────────
    path('borrow/confirm/<int:request_id>/', views.staff_confirm_borrow, name='staff_confirm_borrow'),
    path('return/<int:transaction_id>/',     views.return_item,          name='return_item'),
    path('transaction/<int:transaction_id>/condition/',      views.update_condition,        name='update_condition'),
    path('transaction/<int:transaction_id>/returned-qty/',   views.update_returned_qty,     name='update_returned_qty'),

    # ── Per-device return (new) ───────────────────────────────────────────────
    path('transaction/<int:transaction_id>/devices/',          views.transaction_devices_json, name='transaction_devices_json'),
    path('transaction/<int:transaction_id>/return-devices/',   views.return_devices,           name='return_devices'),

    # ── Borrow requests ───────────────────────────────────────────────────────
    path('requests/',                          views.borrow_requests,  name='borrow_requests'),
    path('requests/<int:request_id>/decline/', views.decline_request,  name='decline_request'),

    # ── Borrow management ─────────────────────────────────────────────────────
    path('borrow-management/',        views.borrow_management,        name='borrow_management'),
    path('borrow-management/export/', views.export_borrow_management, name='export_borrow_management'),

    # ── Device monitoring ─────────────────────────────────────────────────────
    path('device-monitoring/',                     views.device_monitoring,        name='device_monitoring'),
    path('device-monitoring/save/',                views.device_monitoring_save,   name='device_monitoring_save'),
    path('device-monitoring/<int:row_id>/delete/', views.device_monitoring_delete, name='device_monitoring_delete'),
    path('device-monitoring/export/',              views.export_device_monitoring, name='export_device_monitoring'),
]