from django.shortcuts import redirect
from django.conf import settings


class LoginRequiredMiddleware:
    """
    Redirige al login si el usuario no tiene sesión activa.
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
        ruta = request.path_info
        # print(f">>> RUTA: {ruta} | SESSION usuario: {request.session.get('usuario')} | SESSION KEY: {request.session.session_key}")
        
        es_publica = any(ruta.startswith(r) for r in self.RUTAS_PUBLICAS)

        # Ahora usa el sistema nativo de Django
        if not es_publica and not request.user.is_authenticated:
            return redirect(f"{settings.LOGIN_URL}?next={ruta}")

        return self.get_response(request)