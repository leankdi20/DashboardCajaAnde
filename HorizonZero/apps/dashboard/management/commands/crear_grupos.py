"""
Script para crear grupos con sus permisos en HorizonZero.
Ejecutar con: python manage.py shell < apps/dashboard/management/crear_grupos.py
O copiar y pegar en: python manage.py shell
"""
from django.core.management.base import BaseCommand
from django.contrib.auth.models import Group, Permission
from django.contrib.contenttypes.models import ContentType


class Command(BaseCommand):
    help = 'Crea los grupos de HorizonZero con sus permisos'

    def handle(self, *args, **options):

        def get_perm(codename):
            """Obtiene un permiso por codename."""
            try:
                return Permission.objects.get(codename=codename)
            except Permission.DoesNotExist:
                print(f"  ⚠️  Permiso no encontrado: {codename}")
                return None


        def crear_grupo(nombre, codenames):
            """Crea o actualiza un grupo con los permisos indicados."""
            grupo, creado = Group.objects.get_or_create(name=nombre)
            accion = "creado" if creado else "actualizado"

            permisos = []
            for cod in codenames:
                p = get_perm(cod)
                if p:
                    permisos.append(p)

            grupo.permissions.set(permisos)
            print(f"  ✅ Grupo '{nombre}' {accion} con {len(permisos)} permisos")
            return grupo


        print("\n🔐 Creando grupos de HorizonZero...\n")

        # ── Permisos base (todos los grupos los necesitan) ──
        BASE = [
            "view_dashboard",
            "view_formularios",
            "view_encuestas",
        ]

        # ── Todas las encuestas ──
        TODAS_ENCUESTAS = [
            "view_encuesta_satisfaccion",
            "view_encuesta_satisfaccion_oficina",
            "view_encuesta_experiencia_web",
            "view_encuesta_satisfaccion_whatsapp_agente",
            "view_encuesta_satisfaccion_whatsapp",
            "view_encuesta_feria_salud",
        ]

        # ── Todos los formularios ──
        TODOS_FORMULARIOS = [
            # Tarjetas
            "view_formulario_tarjetas",
            "view_soli_tarj_credito",
            "view_soli_tarj_debito",
            "view_soli_tarj_debito_gestion",
            "view_soli_redencion_puntos",
            "view_caja_ande_asistencia",
            # Ahorros
            "view_formulario_ahorros",
            "view_soli_deposito_salario",
            "view_soli_ahorro_mod_cuota",
            "view_soli_reinversion_ahorro",
            "view_soli_autorizacion_ahorro_nuevo",
            # Vivienda
            "view_formulario_vivienda",
            "view_soli_compra_vehiculo",
            "view_soli_prestamo_vivienda",
            "view_soli_prestamo_desarrollo",
            # Préstamos
            "view_formulario_prestamos",
            "view_soli_presolicitud_credito",
            # Control Crédito
            "view_formulario_control_credito",
            "view_comprobante_autorizacion_ahorro",
            "view_comprobantes_pago",
            # Servicio Accionista
            "view_formulario_servicio_accionista",
            "view_soli_clave_cajatel",
            # Seguros
            "view_formulario_seguros",
            "view_soli_seguro_viajero",
            "view_soli_marchamo",
        ]

        # ════════════════════════════════════════════════
        # GRUPOS
        # ════════════════════════════════════════════════

        # 1. Administrador — acceso total
        crear_grupo("Administrador", BASE + TODAS_ENCUESTAS + TODOS_FORMULARIOS)

        # 2. Analista General — ve todo pero no puede administrar
        crear_grupo("Analista General", BASE + TODAS_ENCUESTAS + TODOS_FORMULARIOS)

        # 3. Solo Encuestas
        crear_grupo("Analista Encuestas", BASE + TODAS_ENCUESTAS)

        # 4. Analista Tarjetas
        crear_grupo("Analista Tarjetas", BASE + [
            "view_formulario_tarjetas",
            "view_soli_tarj_credito",
            "view_soli_tarj_debito",
            "view_soli_tarj_debito_gestion",
            "view_soli_redencion_puntos",
            "view_caja_ande_asistencia",
        ])

        # 5. Analista Ahorros
        crear_grupo("Analista Ahorros", BASE + [
            "view_formulario_ahorros",
            "view_soli_deposito_salario",
            "view_soli_ahorro_mod_cuota",
            "view_soli_reinversion_ahorro",
            "view_soli_autorizacion_ahorro_nuevo",
        ])

        # 6. Analista Vivienda y Préstamos
        crear_grupo("Analista Vivienda y Préstamos", BASE + [
            "view_formulario_vivienda",
            "view_soli_compra_vehiculo",
            "view_soli_prestamo_vivienda",
            "view_soli_prestamo_desarrollo",
            "view_formulario_prestamos",
            "view_soli_presolicitud_credito",
        ])

        # 7. Analista Control de Crédito
        crear_grupo("Analista Control de Crédito", BASE + [
            "view_formulario_control_credito",
            "view_comprobante_autorizacion_ahorro",
            "view_comprobantes_pago",
        ])

        # 8. Analista Seguros y Servicio al Accionista
        crear_grupo("Analista Seguros", BASE + [
            "view_formulario_seguros",
            "view_soli_seguro_viajero",
            "view_soli_marchamo",
            "view_formulario_servicio_accionista",
            "view_soli_clave_cajatel",
        ])

        print("\n✅ Grupos creados exitosamente.\n")
        print("📋 Resumen de grupos disponibles:")
        for g in Group.objects.all().order_by("name"):
            print(f"   • {g.name} ({g.permissions.count()} permisos)")

        self.stdout.write(self.style.SUCCESS('Grupos creados exitosamente'))