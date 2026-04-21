import logging
import requests
from django.contrib.auth import get_user_model
from django.conf import settings
from .services.auth_api_service_ldap import AuthApiServiceLDAP
from apps.core.audit import audit_login_fail

logger = logging.getLogger(__name__)
User = get_user_model()


class LDAPAuthBackend:

    def authenticate(self, request, username=None, password=None, **kwargs):
        if not username or not password:
            return None

        ip = ""
        if request:
            ip = request.META.get("HTTP_X_FORWARDED_FOR", "")
            if ip:
                ip = ip.split(",")[0].strip()
            else:
                ip = request.META.get("REMOTE_ADDR", "")

        # Paso 1: ¿Existe en BD local y está autorizado?
        try:
            user = User.objects.get(username=username, is_active=True)
        except User.DoesNotExist:
            logger.warning("Usuario no autorizado intentó ingresar: '%s'", username)
            audit_login_fail(
                username_intentado=username,
                ip=ip,
                descripcion=f"Usuario '{username}' no existe o no está autorizado en el sistema",
            )
            return None

        # Paso 2: Validar contraseña contra LDAP
        auth_service = AuthApiServiceLDAP(
            base_url=settings.AUTH_API_LDAP_URL,
            timeout=15,
        )

        login_data = auth_service.validate_user(username, password)

        if login_data is None:
            logger.warning("Credenciales LDAP inválidas para: '%s'", username)
            audit_login_fail(
                username_intentado=username,
                ip=ip,
                descripcion=f"Contraseña incorrecta para el usuario '{username}'",
            )
            return None

        logger.info("Login exitoso: '%s'", username)
        return user

    def get_user(self, user_id):
        try:
            return User.objects.get(pk=user_id, is_active=True)
        except User.DoesNotExist:
            return None