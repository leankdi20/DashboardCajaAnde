from functools import wraps
from django.shortcuts import redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required, permission_required
from django.core.exceptions import PermissionDenied


def permiso_requerido(permiso: str, redirigir_a: str = "dashboard:home"):
    """
        Verifica un permiso específico.
        Uso @permiso_requerido("dashboard.view_encuestas")
    """

    def decorator(view_func):
        @wraps(view_func)
        def wrapper(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect("usuarios:login")
            if not request.user.has_perm(permiso):
                messages.error(request, "No tienes permiso para acceder a esta página.")
                return redirect(redirigir_a)
            return view_func(request, *args, **kwargs)
        return wrapper
    return decorator

def dashboard_permission(codename):
    """
    Decorador que verifica un permiso específico del dashboard.
    Redirige a 403 si no tiene permiso.
 
    Uso:
        @dashboard_permission("view_soli_tarj_credito")
        def mi_vista(request):
            ...
    """
    def decorator(view_func):
        @wraps(view_func)
        @login_required
        def wrapper(request, *args, **kwargs):
            perm = f"dashboard.{codename}"
            if not request.user.has_perm(perm):
                raise PermissionDenied
            return view_func(request, *args, **kwargs)
        return wrapper
    return decorator
 
 
def tiene_permiso(user, codename):
    """Verifica si un usuario tiene un permiso del dashboard."""
    return user.has_perm(f"dashboard.{codename}")
 