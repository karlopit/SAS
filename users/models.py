from django.contrib.auth.models import AbstractUser
from django.db import models


class CustomUser(AbstractUser):
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('staff', 'Staff'),
    ]

    role = models.CharField(
        max_length=20,
        choices=ROLE_CHOICES,
        default='staff'
    )

    middle_initial = models.CharField(max_length=5, blank=True)

    REQUIRED_FIELDS = []

    def save(self, *args, **kwargs):
        # 🔥 AUTO-FIX ROLE FOR SUPERUSER
        if self.is_superuser:
            self.role = 'admin'
        elif not self.role:
            self.role = 'staff'

        super().save(*args, **kwargs)

    def get_full_name(self):
        first = (self.first_name or "").strip()
        last  = (self.last_name or "").strip()
        mi    = (self.middle_initial or "").strip()

        if not first and not last:
            return self.username

        if first and mi and last:
            return f"{first} {mi}. {last}"
        if first and last:
            return f"{first} {last}"
        if first:
            return first
        return last

    @property
    def display_name(self):
        return self.get_full_name()

    def __str__(self):
        return f"{self.get_full_name()} ({self.role})"