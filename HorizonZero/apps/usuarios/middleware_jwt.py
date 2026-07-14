# ═══════════════════════════════════════════════════════════════════
# apps/usuarios/middleware_jwt.py
# Reconstruye request.user desde el JWT almacenado en cookie
# ═══════════════════════════════════════════════════════════════════
import logging
import requests as http_requests
from django.conf import settings
from django.contrib.auth.models import AnonymousUser

from .session_control import SESSION_REPLACED_REASON

logger = logging.getLogger(__name__)


class UsuarioJWT:
    """
    Objeto que simula django.contrib.auth.models.User
    pero construido desde los datos del JWT.
    No hace ninguna consulta a la BD.
    """

    def __init__(self, datos: dict):
        self.id              = datos.get("id")
        self.pk              = datos.get("id")
        self.username        = datos.get("username", "")
        self.first_name      = datos.get("first_name", "")
        self.last_name       = datos.get("last_name", "")
        self.email           = datos.get("email", "")
        self.is_staff        = datos.get("is_staff", False)
        self.is_active       = True
        self.is_authenticated = True
        self.is_anonymous    = False
        self._permisos       = set(datos.get("permisos", []))
        self._grupos         = datos.get("grupos", [])
        self._perfil         = datos.get("perfil", {})

    def has_perm(self, perm, obj=None):
        if self.is_staff:
            return True
        return perm in self._permisos

    def has_perms(self, perms, obj=None):
        return all(self.has_perm(p) for p in perms)

    def has_module_perms(self, app_label):
        if self.is_staff:
            return True
        return any(p.startswith(f"{app_label}.") for p in self._permisos)

    def get_full_name(self):
        return f"{self.first_name} {self.last_name}".strip() or self.username

    def get_short_name(self):
        return self.first_name or self.username

    @property
    def groups(self):
        return self._grupos

    @property
    def perfil(self):
        return type("Perfil", (), {
            "unidad":   self._perfil.get("unidad", ""),
            "sucursal": self._perfil.get("sucursal", ""),
        })()


class JWTAuthMiddleware:
    """
    Lee el token JWT de la cookie 'hz_token',
    verifica con la API y reconstruye request.user.
    """

    COOKIE_NAME = "hz_token"

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        request.session_invalidated = False
        request.session_invalid_reason = None
        token = request.COOKIES.get(self.COOKIE_NAME)

        if token:
            try:
                api_url  = settings.API_URL
                response = http_requests.get(
                    f"{api_url}/api/auth/me/",
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Connection": "close",
                        "Accept": "application/json",
                    },
                    timeout=10,
                    proxies={"http": None, "https": None},
                )
                if response.status_code == 200:
                    datos = response.json()
                    request.user     = UsuarioJWT(datos)
                    request.jwt_token = token
                elif response.status_code in (401, 403):
                    request.session_invalidated = True
                    request.session_invalid_reason = SESSION_REPLACED_REASON
                    request.user = AnonymousUser()
                else:
                    request.user = AnonymousUser()
            except Exception as e:
                logger.warning(f"[JWTAuthMiddleware] Error verificando token: {e}")
                request.user = AnonymousUser()
        else:
            request.user = AnonymousUser()

        return self.get_response(request)
