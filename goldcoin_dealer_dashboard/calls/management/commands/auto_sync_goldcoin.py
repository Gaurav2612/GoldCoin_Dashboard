
from django.core.management.base import BaseCommand, CommandError
from django.utils import timezone

from calls.views import sync_goldcoin_excel


class Command(BaseCommand):

    help = "Automatically synchronize the latest Gold Coin Excel file."

    def handle(self, *args, **options):

        start_time = timezone.now()

        self.stdout.write(
            self.style.WARNING(
                "\nStarting Gold Coin Excel synchronization..."
            )
        )

        try:
            result = sync_goldcoin_excel()

            self.stdout.write(
                self.style.SUCCESS(
                    "\nExcel synchronization completed successfully."
                )
            )

            self.stdout.write(str(result))

        except Exception as exc:

            self.stderr.write(
                self.style.ERROR(
                    f"\nExcel synchronization failed: {exc}"
                )
            )

            raise CommandError(str(exc)) from exc

        finally:

            duration = (
                timezone.now() - start_time
            ).total_seconds()

            self.stdout.write(
                f"\nFinished in {duration:.2f} seconds."
            )
