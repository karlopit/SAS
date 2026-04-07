from django.test import TestCase
from users.models import CustomUser
from .models import Item, Transaction

class ItemTestCase(TestCase):
    def setUp(self):
        self.item = Item.objects.create(name='Laptop', quantity=5, available_quantity=5)

    def test_item_created(self):
        self.assertEqual(self.item.name, 'Laptop')

    def test_available_quantity(self):
        self.assertEqual(self.item.available_quantity, 5)

class TransactionTestCase(TestCase):
    def setUp(self):
        self.user = CustomUser.objects.create_user(username='borrower1', password='pass1234')
        self.item = Item.objects.create(name='Camera', quantity=3, available_quantity=3)

    def test_borrow_reduces_availability(self):
        self.item.available_quantity -= 1
        self.item.save()
        self.assertEqual(self.item.available_quantity, 2)