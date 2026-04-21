# ═══════════════════════════════════════════════════════════════════
# apps/core/audit.py
# Helper centralizado para registrar eventos de auditoría
# ═══════════════════════════════════════════════════════════════════

import json
import logging
from django.utils import timezone

logger = logging.getLogger(__name__)


def _get_ip(request) -> str:
    """Extrae la IP real considerando proxies y load balancers."""
    ip = request.META.get("HTTP_X_FORWARDED_FOR")
    if ip:
        return ip.split(",")[0].strip()
    return request.META.get("REMOTE_ADDR", "")


def _get_user_agent(request) -> str:
    return request.META.get("HTTP_USER_AGENT", "")[:500]


def audit(
    request,
    accion: str,
    modulo: str,
    descripcion: str,
    objeto_id=None,
    objeto_nombre: str = None,
    datos_anteriores: dict = None,
    datos_nuevos: dict = None,
    severidad: str = "INFO",
) -> None:
    """
    Registra una acción de auditoría desde cualquier view autenticada.

    Args:
        request:          HttpRequest con request.user autenticado
        accion:           AuditLog.Accion.XXX
        modulo:           AuditLog.Modulo.XXX
        descripcion:      Texto legible para el analista de la USI
        objeto_id:        ID del registro afectado (opcional)
        objeto_nombre:    Nombre legible del objeto (opcional)
        datos_anteriores: Dict con estado ANTES del cambio (solo UPDATE)
        datos_nuevos:     Dict con estado DESPUÉS del cambio (solo UPDATE/CREATE)
        severidad:        AuditLog.Severidad.INFO / WARNING / CRITICAL
    """
    try:
        from apps.dashboard.models import AuditLog

        usuario  = request.user if request.user.is_authenticated else None
        username = request.user.username if request.user.is_authenticated else "anónimo"

        AuditLog.objects.create(
            usuario          = usuario,
            username         = username,
            accion           = accion,
            modulo           = modulo,
            severidad        = severidad,
            descripcion      = descripcion,
            objeto_id        = str(objeto_id) if objeto_id is not None else None,
            objeto_nombre    = objeto_nombre,
            datos_anteriores = json.dumps(datos_anteriores, ensure_ascii=False) if datos_anteriores else None,
            datos_nuevos     = json.dumps(datos_nuevos,     ensure_ascii=False) if datos_nuevos     else None,
            ip_address       = _get_ip(request) or None,
            user_agent       = _get_user_agent(request),
            url              = request.path,
            metodo_http      = request.method,
        )

    except Exception as e:
        # El log NUNCA debe interrumpir el flujo de la aplicación
        logger.error(f"[AUDIT] Error al registrar evento: {e}")


def audit_login_fail(username_intentado: str, ip: str, descripcion: str = None) -> None:
    try:
        from apps.dashboard.models import AuditLog

        AuditLog.objects.create(
            usuario       = None,
            username      = username_intentado or "desconocido",
            accion        = AuditLog.Accion.LOGIN_FAIL,
            modulo        = AuditLog.Modulo.SISTEMA,
            severidad     = AuditLog.Severidad.CRITICAL,
            descripcion   = descripcion or f"Intento de acceso fallido para '{username_intentado}'",
            ip_address    = ip or None,
        )
    except Exception as e:
        logger.error(f"[AUDIT] Error al registrar login fallido: {e}")