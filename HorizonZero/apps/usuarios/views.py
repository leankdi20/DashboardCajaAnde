import logging
import requests
from django.shortcuts import render, redirect
from django.contrib import messages, auth
from django.contrib.auth import authenticate, login, logout
from django.views.decorators.http import require_http_methods
from django.views.decorators.cache import never_cache

from .forms import LoginForm
from apps.core.audit import audit
from apps.dashboard.models import AuditLog


logger = logging.getLogger(__name__)


@never_cache
@require_http_methods(["GET", "POST"])
def login_view(request):
    # Si ya está autenticado, redirigir
    if request.user.is_authenticated:
        return redirect("dashboard:home")

    if request.method == "GET":
        return render(request, "login/login.html", {"form": LoginForm()})

    form = LoginForm(request.POST)
    if not form.is_valid():
        return render(request, "login/login.html", {"form": form})

    username = form.cleaned_data["usuario"]
    password = form.cleaned_data["password"]

    try:
        user = authenticate(request, username=username, password=password)
    except requests.exceptions.ConnectionError:
        messages.error(request, "No se pudo conectar con el servidor de autenticación.")
        return render(request, "login/login.html", {"form": form})
    except requests.exceptions.Timeout:
        messages.error(request, "El servidor de autenticación tardó demasiado.")
        return render(request, "login/login.html", {"form": form})
    except requests.exceptions.RequestException:
        messages.error(request, "Error inesperado. Intente más tarde.")
        return render(request, "login/login.html", {"form": form})

    if user is None:
        messages.error(request, "Usuario no autorizado o credenciales incorrectas.")
        return render(request, "login/login.html", {"form": form})

    # Login nativo de Django — maneja session fixation automáticamente
    login(request, user, backend="apps.usuarios.backends.LDAPAuthBackend")
    logger.info("Usuario '%s' inició sesión.", username)


    # ── Auditoría: login exitoso ──────────────────────────────
    audit(request, AuditLog.Accion.LOGIN_OK, AuditLog.Modulo.SISTEMA,
          f"Inicio de sesión exitoso — {user.get_full_name() or username}",
          severidad=AuditLog.Severidad.INFO)


    next_url = request.GET.get("next", "")
    if next_url and next_url.startswith("/") and not next_url.startswith("//"):
        return redirect(next_url)

    return redirect("dashboard:home")


@never_cache
def logout_view(request):
    logger.info("Usuario '%s' cerró sesión.", request.user.username)

    # ── Auditoría: logout — antes de cerrar la sesión ─────────
    audit(request, AuditLog.Accion.LOGOUT, AuditLog.Modulo.SISTEMA,
          f"Cierre de sesión — {request.user.get_full_name() or request.user.username}",
          severidad=AuditLog.Severidad.INFO)

    logout(request)
    return redirect("usuarios:login")