from django.urls import path
from . import views

urlpatterns = [
    path('', views.welcome, name='welcome'),
    path('dashboard/', views.index, name='index'),
    path('add/', views.add_item, name='add_item'),
    path('item/<int:item_id>/edit/', views.edit_item, name='edit_item'),  # ADD THIS LINE
    path('borrow/confirm/<int:request_id>/', views.staff_confirm_borrow, name='staff_confirm_borrow'),
    path('return/<int:transaction_id>/', views.return_item, name='return_item'),
    path('transaction/<int:transaction_id>/condition/', views.update_condition, name='update_condition'),
    path('transaction/<int:transaction_id>/returned-qty/', views.update_returned_qty, name='update_returned_qty'),
    path('requests/', views.borrow_requests, name='borrow_requests'),
    path('requests/<int:request_id>/decline/', views.decline_request, name='decline_request'),
    path('borrow-management/', views.borrow_management, name='borrow_management'),
    path('borrow-management/export/', views.export_borrow_management, name='export_borrow_management'),
    path('device-monitoring/', views.device_monitoring, name='device_monitoring'),
    path('device-monitoring/save/', views.device_monitoring_save, name='device_monitoring_save'),
    path('device-monitoring/<int:row_id>/delete/', views.device_monitoring_delete, name='device_monitoring_delete'),
    path('device-monitoring/export/', views.export_device_monitoring, name='export_device_monitoring'),
]