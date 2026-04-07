from django.test import TestCase
from .models import CustomUser

class CustomUserTestCase(TestCase):
    def setUp(self):
        self.user = CustomUser.objects.create_user(
            username='testuser', password='pass1234', role='staff'
        )

    def test_user_created(self):
        self.assertEqual(self.user.username, 'testuser')

    def test_user_role(self):
        self.assertEqual(self.user.role, 'staff')