from django.shortcuts import redirect
from django.conf import settings

from .session_control import (
    APISessionExpiredError,
    SESSION_INVALID_REASON,
    build_forced_logout_response,
    is_async_request,
)


class ForcedLogoutMiddleware:
    """
    Centraliza el cierre de sesion cuando la API invalida el token actual.
    """

    RUTAS_PUBLICAS = [
        "/usuarios/login/",
        "/usuarios/logout/",
        "/admin/",
        "/static/",
    ]

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        try:
            if self._should_force_logout(request):
                return build_forced_logout_response(
                    request,
                    reason=getattr(request, "session_invalid_reason", None),
                )
            return self.get_response(request)
        except APISessionExpiredError as exc:
            return build_forced_logout_response(
                request,
                reason=exc.reason,
                message=exc.message,
            )

    def _should_force_logout(self, request):
        ruta = request.path_info
        es_publica = any(ruta.startswith(r) for r in self.RUTAS_PUBLICAS)
        return not es_publica and getattr(request, "session_invalidated", False)


class LoginRequiredMiddleware:
    """
    Redirige al login si el usuario no tiene sesión activa.
    """
    RUTAS_PUBLICAS = [
        "/usuarios/login/",  
        "/usuarios/logout/",
        "/usuarios/session-status/",
        "/admin/",            
        "/static/",
    ]

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        ruta = request.path_info
        # print(f">>> RUTA: {ruta} | SESSION usuario: {request.session.get('usuario')} | SESSION KEY: {request.session.session_key}")
        
        es_publica = any(ruta.startswith(r) for r in self.RUTAS_PUBLICAS)

        # Ahora usa el sistema nativo de Django
        if not es_publica and not request.user.is_authenticated:
            if is_async_request(request):
                return build_forced_logout_response(
                    request,
                    reason=SESSION_INVALID_REASON,
                    status=401,
                )
            return redirect(f"{settings.LOGIN_URL}?next={ruta}")

        return self.get_response(request)
