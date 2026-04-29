import os
from django.core.management import call_command
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = 'Flush the entire database IF DJANGO_FLUSH_DATABASE=true'

    def handle(self, *args, **options):
        if os.environ.get('DJANGO_FLUSH_DATABASE', '').lower() == 'true':
            self.stdout.write('Flushing database…')
            # --no-input avoids the confirmation prompt
            call_command('flush', interactive=False, verbosity=1)
            self.stdout.write(self.style.SUCCESS('Database flushed successfully.'))
        else:
            self.stdout.write('SKIPPED: DJANGO_FLUSH_DATABASE is not "true".')