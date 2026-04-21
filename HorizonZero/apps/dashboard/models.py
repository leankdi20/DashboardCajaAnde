from django.db import models

# Create your models here.



class DashboardAccess(models.Model):
    class Meta:
        managed = False
        default_permissions = ()
        permissions = [
 
            # ─────────────────────────────────────────────
            # MÓDULOS PRINCIPALES
            # ─────────────────────────────────────────────
            ("view_dashboard",   "Puede acceder al dashboard"),
            ("view_encuestas",   "Puede ver módulo Encuestas"),
            ("view_formularios", "Puede ver módulo Formularios"),
            # ─────────────────────────────────────────────
            # USUARIOS
            # ─────────────────────────────────────────────
            ("view_usuarios",    "Puede ver módulo Usuarios"),    
            # ─────────────────────────────────────────────
            # LOGS
            # ─────────────────────────────────────────────
            ("view_logs",        "Puede ver módulo Logs"),   



            # ─────────────────────────────────────────────
            # CRUD USUARIOS
            # ─────────────────────────────────────────────
            ("view_agentes",   "Agentes: Ver listado"),
            ("add_agente",     "Agentes: Crear agente"),
            ("change_agente",  "Agentes: Editar agente"),
            ("delete_agente",  "Agentes: Eliminar agente"),


 
            # ─────────────────────────────────────────────
            # ENCUESTAS
            # ─────────────────────────────────────────────
            ("view_encuesta_satisfaccion",                "Encuesta: Satisfacción"),
            ("view_encuesta_satisfaccion_oficina",         "Encuesta: Satisfacción Oficina Digital"),
            ("view_encuesta_experiencia_web",              "Encuesta: Experiencia en página web"),
            ("view_encuesta_satisfaccion_whatsapp_agente", "Encuesta: Satisfacción WhatsApp Agente"),
            ("view_encuesta_satisfaccion_whatsapp",        "Encuesta: Satisfacción WhatsApp"),
            ("view_encuesta_feria_salud",                  "Encuesta: Feria de la Salud"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — TARJETAS
            # ─────────────────────────────────────────────
            ("view_formulario_tarjetas",              "Formularios: Tarjetas (categoría)"),
            ("view_soli_tarj_credito",                "Tarjetas: Solicitud gestión tarjeta de crédito"),
            ("view_soli_tarj_debito",                 "Tarjetas: Solicitud tarjeta Ciudadano de Oro"),
            ("view_soli_tarj_debito_gestion",         "Tarjetas: Solicitud gestión tarjeta de débito"),
            ("view_soli_redencion_puntos",             "Tarjetas: Solicitud redención de puntos"),
            ("view_caja_ande_asistencia",              "Tarjetas: Caja de ANDE Asistencia"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — AHORROS
            # ─────────────────────────────────────────────
            ("view_formulario_ahorros",               "Formularios: Ahorros (categoría)"),
            ("view_soli_deposito_salario",             "Ahorros: Solicitud depósito de salario"),
            ("view_soli_ahorro_mod_cuota",             "Ahorros: Modificación de cuota"),
            ("view_soli_reinversion_ahorro",           "Ahorros: Reinversión ahorro existente"),
            ("view_soli_autorizacion_ahorro_nuevo",    "Ahorros: Autorización ahorro nuevo"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — VIVIENDA
            # ─────────────────────────────────────────────
            ("view_formulario_vivienda",              "Formularios: Vivienda (categoría)"),
            ("view_soli_compra_vehiculo",              "Vivienda: Solicitud compra de vehículo"),
            ("view_soli_prestamo_vivienda",            "Vivienda: Préstamo para vivienda"),
            ("view_soli_prestamo_desarrollo",          "Vivienda: Préstamo desarrollo económico"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — PRÉSTAMOS
            # ─────────────────────────────────────────────
            ("view_formulario_prestamos",             "Formularios: Préstamos (categoría)"),
            ("view_soli_presolicitud_credito",         "Préstamos: Pre-solicitud crédito personal"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — CONTROL DE CRÉDITO
            # ─────────────────────────────────────────────
            ("view_formulario_control_credito",       "Formularios: Control de Crédito (categoría)"),
            ("view_comprobante_autorizacion_ahorro",   "Control Crédito: Comprobante autorización ahorro"),
            ("view_comprobantes_pago",                 "Control Crédito: Comprobantes pago depósitos"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — SERVICIO AL ACCIONISTA
            # ─────────────────────────────────────────────
            ("view_formulario_servicio_accionista",   "Formularios: Servicio al Accionista (categoría)"),
            ("view_soli_clave_cajatel",                "Servicio Accionista: Solicitud clave CajaTel"),
 
            # ─────────────────────────────────────────────
            # FORMULARIOS — SEGUROS
            # ─────────────────────────────────────────────
            ("view_formulario_seguros",               "Formularios: Seguros (categoría)"),
            ("view_soli_seguro_viajero",               "Seguros: Seguro viajero"),
            ("view_soli_marchamo",                     "Seguros: Seguro marchamo"),



       


        ]


# ═══════════════════════════════════════════════════════════════════
# apps/dashboard/models.py
# Agregar este modelo AL FINAL del archivo, después de DashboardAccess
# ═══════════════════════════════════════════════════════════════════

from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
from datetime import timedelta


class AuditLog(models.Model):
    """
    Registro de auditoría de todas las acciones realizadas en HorizonZero.
    Diseñado para cumplir con los requerimientos de la Unidad de Seguridad
    Institucional (USI) — Caja de ANDE.

    Política de retención: 2 años (730 días).
    Los registros NO pueden ser editados ni eliminados desde la aplicación.
    Solo superusuario puede purgar registros vencidos vía management command.
    """

    # ── Categorías de acción ─────────────────────────────────────
    class Accion(models.TextChoices):
        # Autenticación
        LOGIN_OK         = "LOGIN_OK",         "Inicio de sesión exitoso"
        LOGIN_FAIL       = "LOGIN_FAIL",       "Intento de acceso fallido"
        LOGOUT           = "LOGOUT",           "Cierre de sesión"
        # CRUD Agentes
        AGENTE_CREATE    = "AGENTE_CREATE",    "Agente creado"
        AGENTE_UPDATE    = "AGENTE_UPDATE",    "Agente editado"
        AGENTE_DELETE    = "AGENTE_DELETE",    "Agente eliminado"
        AGENTE_RESTORE   = "AGENTE_RESTORE",   "Agente restaurado"
        # Acceso a módulos
        VIEW_REPORTE     = "VIEW_REPORTE",     "Acceso a reporte"
        VIEW_FORMULARIO  = "VIEW_FORMULARIO",  "Acceso a formulario"
        # Exportaciones
        EXPORT_EXCEL     = "EXPORT_EXCEL",     "Exportación Excel"
        EXPORT_QR        = "EXPORT_QR",        "Descarga QR"
        EXPORT_ZIP       = "EXPORT_ZIP",       "Descarga ZIP QR"
        # Seguridad
        PERMISO_DENEGADO = "PERMISO_DENEGADO", "Acceso denegado por permisos"

    # ── Módulos del sistema ──────────────────────────────────────
    class Modulo(models.TextChoices):
        SISTEMA     = "SISTEMA",      "Sistema"
        AGENTES     = "AGENTES",      "Gestión de Agentes"
        ENCUESTAS   = "ENCUESTAS",    "Encuestas"
        FORMULARIOS = "FORMULARIOS",  "Formularios"
        USUARIOS    = "USUARIOS",     "Usuarios"
        QR          = "QR",           "Códigos QR"

    # ── Severidad del evento ─────────────────────────────────────
    class Severidad(models.TextChoices):
        INFO     = "INFO",     "Informativo"
        WARNING  = "WARNING",  "Advertencia"
        CRITICAL = "CRITICAL", "Crítico"

    # ── Campos ───────────────────────────────────────────────────

    # FK al usuario Django — se pone NULL si el usuario es eliminado
    # pero username (abajo) siempre queda como texto plano
    usuario = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="audit_logs",
        verbose_name="Usuario",
        db_index=True,
    )

    # Texto plano — persiste aunque el usuario sea eliminado del sistema
    username = models.CharField(
        max_length=150,
        verbose_name="Nombre de usuario",
        db_index=True,
    )

    accion = models.CharField(
        max_length=50,
        choices=Accion.choices,
        verbose_name="Acción",
        db_index=True,
    )

    modulo = models.CharField(
        max_length=50,
        choices=Modulo.choices,
        verbose_name="Módulo",
        db_index=True,
    )

    severidad = models.CharField(
        max_length=10,
        choices=Severidad.choices,
        default=Severidad.INFO,
        verbose_name="Severidad",
        db_index=True,
    )

    # Descripción legible para el analista de la USI
    descripcion = models.CharField(
        max_length=500,
        verbose_name="Descripción",
    )

    # ID del objeto afectado (agente_id, respuesta_id, etc.)
    objeto_id = models.CharField(
        max_length=50,
        blank=True,
        null=True,
        verbose_name="ID del objeto",
    )

    # Nombre legible del objeto afectado
    objeto_nombre = models.CharField(
        max_length=255,
        blank=True,
        null=True,
        verbose_name="Nombre del objeto",
    )

    # Estado ANTES del cambio — solo para UPDATE (JSON como texto)
    datos_anteriores = models.TextField(
        blank=True,
        null=True,
        verbose_name="Datos anteriores",
        help_text="Estado del objeto ANTES del cambio (JSON)",
    )

    # Estado DESPUÉS del cambio — solo para UPDATE (JSON como texto)
    datos_nuevos = models.TextField(
        blank=True,
        null=True,
        verbose_name="Datos nuevos",
        help_text="Estado del objeto DESPUÉS del cambio (JSON)",
    )

    # Información de red para trazabilidad
    ip_address = models.GenericIPAddressField(
        blank=True,
        null=True,
        verbose_name="Dirección IP",
        db_index=True,
    )

    user_agent = models.CharField(
        max_length=500,
        blank=True,
        null=True,
        verbose_name="Navegador / Agente",
    )

    url = models.CharField(
        max_length=500,
        blank=True,
        null=True,
        verbose_name="URL accedida",
    )

    metodo_http = models.CharField(
        max_length=10,
        blank=True,
        null=True,
        verbose_name="Método HTTP",
    )

    # Timestamp inmutable
    fecha = models.DateTimeField(
        default=timezone.now,
        verbose_name="Fecha y hora",
        db_index=True,
        editable=False,
    )

    # Política de retención: 2 años desde la fecha del evento
    fecha_vencimiento = models.DateTimeField(
        verbose_name="Vence el",
        db_index=True,
        editable=False,
    )

    # ── Meta ─────────────────────────────────────────────────────
    class Meta:
        app_label           = "dashboard"
        verbose_name        = "Registro de auditoría"
        verbose_name_plural = "Registros de auditoría"
        ordering            = ["-fecha"]
        indexes = [
            models.Index(fields=["username", "fecha"],   name="idx_audit_user_fecha"),
            models.Index(fields=["accion", "fecha"],     name="idx_audit_accion_fecha"),
            models.Index(fields=["modulo", "fecha"],     name="idx_audit_modulo_fecha"),
            models.Index(fields=["severidad", "fecha"],  name="idx_audit_sev_fecha"),
            models.Index(fields=["ip_address"],          name="idx_audit_ip"),
            models.Index(fields=["fecha_vencimiento"],   name="idx_audit_vencimiento"),
        ]

    def save(self, *args, **kwargs):
        # Calcular fecha de vencimiento si no está seteada
        if not self.fecha_vencimiento:
            self.fecha_vencimiento = (self.fecha or timezone.now()) + timedelta(days=730)
        super().save(*args, **kwargs)

    def __str__(self):
        return f"[{self.fecha:%d/%m/%Y %H:%M}] {self.username} — {self.get_accion_display()}"

    @property
    def esta_vencido(self):
        return timezone.now() > self.fecha_vencimiento

    @property
    def severidad_color(self):
        return {
            "INFO":     "blue",
            "WARNING":  "amber",
            "CRITICAL": "red",
        }.get(self.severidad, "slate")


# ═══════════════════════════════════════════════════════════════════
# apps/core/audit.py  — helper para registrar desde cualquier view
# ═══════════════════════════════════════════════════════════════════

"""
USO EN VIEWS:

from apps.core.audit import audit

# Login exitoso
audit(request, AuditLog.Accion.LOGIN_OK, AuditLog.Modulo.SISTEMA,
      "Inicio de sesión exitoso")

# Crear agente
audit(request, AuditLog.Accion.AGENTE_CREATE, AuditLog.Modulo.AGENTES,
      f"Creó el agente '{nombre}'",
      objeto_id=nuevo_id, objeto_nombre=nombre,
      datos_nuevos={"nombre": nombre, "unidad_id": unidad_id, "sucursal_id": sucursal_id})

# Editar agente (con diff antes/después)
audit(request, AuditLog.Accion.AGENTE_UPDATE, AuditLog.Modulo.AGENTES,
      f"Editó el agente '{agente[\"nombre\"]}'",
      objeto_id=agente_id, objeto_nombre=agente["nombre"],
      datos_anteriores={"nombre": agente["nombre"], "unidad_id": agente["unidad_id"]},
      datos_nuevos={"nombre": nombre, "unidad_id": unidad_id})

# Eliminar agente
audit(request, AuditLog.Accion.AGENTE_DELETE, AuditLog.Modulo.AGENTES,
      f"Eliminó el agente '{agente[\"nombre\"]}'",
      objeto_id=agente_id, objeto_nombre=agente["nombre"],
      severidad=AuditLog.Severidad.WARNING)

# Login fallido (sin request.user autenticado)
audit_anonimo(username_intentado, ip, AuditLog.Accion.LOGIN_FAIL,
              "Credenciales incorrectas", severidad=AuditLog.Severidad.CRITICAL)
"""