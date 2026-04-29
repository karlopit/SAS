"""
Idempotently create/update a superuser from environment variables.

Set these env vars in Render:
  DJANGO_SUPERUSER_USERNAME
  DJANGO_SUPERUSER_FULLNAME    (e.g. "Juan M. Dela Cruz")
  DJANGO_SUPERUSER_PASSWORD

The command runs at every deployment and updates the user if any env var changed.
"""
import os
from django.core.management.base import BaseCommand
from users.models import CustomUser


class Command(BaseCommand):
    help = 'Ensures a superuser exists based on environment variables.'

    def handle(self, *args, **options):
        username = os.environ.get('DJANGO_SUPERUSER_USERNAME')
        fullname = os.environ.get('DJANGO_SUPERUSER_FULLNAME', '')
        password = os.environ.get('DJANGO_SUPERUSER_PASSWORD')

        if not username or not password:
            self.stdout.write('Superuser env vars not set – skipping.')
            return

        user, created = CustomUser.objects.get_or_create(username=username)

        # Set superuser flags
        user.is_staff = True
        user.is_superuser = True

        # Use the full name string as first_name so get_full_name() displays it
        if fullname:
            user.first_name = fullname
            user.last_name = ''   # keep empty so nothing extra is appended

        user.set_password(password)
        user.save()

        action = 'Created' if created else 'Updated'
        self.stdout.write(f'{action} superuser “{username}” (name: {fullname}).')