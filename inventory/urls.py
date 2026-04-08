from django.urls import path
from . import views

urlpatterns = [
    path('', views.welcome, name='welcome'),
    path('dashboard/', views.index, name='index'),
    path('add/', views.add_item, name='add_item'),
    path('borrow/confirm/<int:request_id>/', views.staff_confirm_borrow, name='staff_confirm_borrow'),
    path('return/<int:transaction_id>/', views.return_item, name='return_item'),
    path('transaction/<int:transaction_id>/condition/', views.update_condition, name='update_condition'),
    path('requests/', views.borrow_requests, name='borrow_requests'),
    path('requests/<int:request_id>/decline/', views.decline_request, name='decline_request'),
]