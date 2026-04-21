# ═══════════════════════════════════════════════════════════════════
# apps/dashboard/management/commands/purgar_audit_logs.py
#
# Comando para purgar registros de auditoría vencidos (> 2 años).
# Solo puede ejecutarlo un superusuario desde el servidor.
#
# Uso:
#   python manage.py purgar_audit_logs
#   python manage.py purgar_audit_logs --dry-run   (solo muestra cuántos borraría)
#
# Recomendado: programar con Task Scheduler (Windows) o cron (Linux)
# una vez por mes.
# ═══════════════════════════════════════════════════════════════════

from django.core.management.base import BaseCommand
from django.utils import timezone
from apps.dashboard.models import AuditLog


class Command(BaseCommand):
    help = "Purga registros de auditoría vencidos (política de retención: 2 años)"

    def add_arguments(self, parser):
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Muestra cuántos registros se eliminarían sin borrarlos",
        )

    def handle(self, *args, **options):
        ahora     = timezone.now()
        vencidos  = AuditLog.objects.filter(fecha_vencimiento__lt=ahora)
        cantidad  = vencidos.count()

        if options["dry_run"]:
            self.stdout.write(
                self.style.WARNING(
                    f"[DRY RUN] Se eliminarían {cantidad} registros vencidos."
                )
            )
            return

        if cantidad == 0:
            self.stdout.write(self.style.SUCCESS("No hay registros vencidos para purgar."))
            return

        vencidos.delete()
        self.stdout.write(
            self.style.SUCCESS(
                f"✓ Purgados {cantidad} registros de auditoría vencidos al {ahora:%d/%m/%Y %H:%M}."
            )
        )