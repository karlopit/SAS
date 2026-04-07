from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('add/', views.add_item, name='add_item'),
    path('borrow/', views.borrow_item, name='borrow_item'),
    path('return/<int:transaction_id>/', views.return_item, name='return_item'),
    path('transaction/<int:transaction_id>/condition/', views.update_condition, name='update_condition'),
]