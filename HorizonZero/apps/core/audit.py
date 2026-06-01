# ═══════════════════════════════════════════════════════════════════
# apps/core/audit.py
# Registra eventos de auditoría via API interna
# ═══════════════════════════════════════════════════════════════════
import logging
import requests
from django.conf import settings

logger = logging.getLogger(__name__)


def _get_ip(request) -> str:
    for header in ["HTTP_X_FORWARDED_FOR", "HTTP_X_REAL_IP", "HTTP_CLIENT_IP"]:
        ip = request.META.get(header)
        if ip:
            return ip.split(",")[0].strip().split(":")[0].strip()
    return request.META.get("REMOTE_ADDR", "")


def _get_user_agent(request) -> str:
    return request.META.get("HTTP_USER_AGENT", "")[:500]


def _post_audit(data: dict) -> None:
    try:
        api_url  = settings.API_URL
        key      = settings.API_INTERNAL_KEY
        print(f">>> AUDIT POST: {data.get('accion')} | URL: {api_url} | KEY: {repr(key)}")
        response = requests.post(
            f"{api_url}/api/logs/crear/",
            json=data,
            headers={
                "Content-Type":   "application/json",
                "X-Internal-Key": key,
            },
            timeout=5,
        )
        print(f">>> AUDIT RESPONSE: {response.status_code} | {response.text}")
    except Exception as e:
        print(f">>> AUDIT ERROR: {e}")
        logger.error(f"[AUDIT] Error enviando a API: {e}")


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
    try:
        username  = request.user.username if request.user.is_authenticated else "anónimo"
        usuario_id = request.user.id if request.user.is_authenticated else None

        _post_audit({
            "username":         username,
            "usuario_id":       usuario_id,
            "accion":           accion,
            "modulo":           modulo,
            "severidad":        severidad,
            "descripcion":      descripcion,
            "objeto_id":        str(objeto_id) if objeto_id is not None else None,
            "objeto_nombre":    objeto_nombre,
            "ip_address":       _get_ip(request) or None,
            "user_agent":       _get_user_agent(request),
            "url":              request.path,
            "metodo_http":      request.method,
        })
    except Exception as e:
        logger.error(f"[AUDIT] Error: {e}")


def audit_login_fail(username_intentado: str, ip: str, descripcion: str = None) -> None:
    try:
        _post_audit({
            "username":    username_intentado or "desconocido",
            "accion":      "LOGIN_FAIL",
            "modulo":      "SISTEMA",
            "severidad":   "CRITICAL",
            "descripcion": descripcion or f"Intento de acceso fallido para '{username_intentado}'",
            "ip_address":  ip or None,
        })
    except Exception as e:
        logger.error(f"[AUDIT] Error login fail: {e}")