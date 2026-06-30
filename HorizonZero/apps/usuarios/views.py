import logging
import requests as http_requests
from django.conf import settings
from django.shortcuts import render, redirect
from django.contrib import messages
from django.views.decorators.http import require_http_methods
from django.views.decorators.cache import never_cache

import time

from .forms import LoginForm
from apps.core.audit import _post_audit
from apps.usuarios.services.login_attempts import (
    clear_failed_attempts,
    get_login_block_remaining_minutes,
    get_login_block_remaining_seconds,
    is_login_blocked,
    register_failed_attempt,
)

logger = logging.getLogger(__name__)


@never_cache
@require_http_methods(["GET", "POST"])
def login_view(request):
    if request.user.is_authenticated:
        return redirect("dashboard:home")

    form = LoginForm(request.POST or None)

    if request.method == "GET":
        return render(request, "login/login.html", {"form": form})

    if not form.is_valid():
        return render(request, "login/login.html", {"form": form})

    username = form.cleaned_data["usuario"].strip().lower()
    password = form.cleaned_data["password"]

    if is_login_blocked(username, request):
        remaining_seconds = get_login_block_remaining_seconds(username, request)
        remaining_minutes = get_login_block_remaining_minutes(username, request)
        messages.error(
            request,
            f"Demasiados intentos fallidos. Por seguridad, espere {remaining_minutes} minuto(s)."
        )
        return render(request, "login/login.html", {
            "form": form,
            "block_remaining_seconds": remaining_seconds,
        })

    # Llamar a la API en lugar de authenticate()
    try:
        api_url  = settings.API_URL
        inicio = time.time()
        response = http_requests.post(
            f"{api_url}/api/auth/login/",
            json={"username": username, "password": password},
            timeout=15,
        )
        print(" LOGIN API LDAP TARDA ESTO: ", round(time.time() - inicio, 2) )

    except http_requests.exceptions.ConnectionError:
        messages.error(request, "No se pudo conectar con el servidor de autenticación.")
        return render(request, "login/login.html", {"form": form})
    except http_requests.exceptions.Timeout:
        messages.error(request, "El servidor de autenticación tardó demasiado.")
        return render(request, "login/login.html", {"form": form})

    if response.status_code == 401:
        register_failed_attempt(username, request)
        # Audit login fallido
        try:
            inicio_audit = time.time()
            _post_audit({
                "username":    username,
                "accion":      "LOGIN_FAIL",
                "modulo":      "SISTEMA",
                "severidad":   "CRITICAL",
                "descripcion": f"Credenciales incorrectas para '{username}'",
                "ip_address":  request.META.get("REMOTE_ADDR", ""),
                "user_agent":  request.META.get("HTTP_USER_AGENT", "")[:500],
                "url":         request.path,
                "metodo_http": request.method,
            })
            print(">>> AUDIT LOGIN_OK tardó:", round(time.time() - inicio_audit, 2))
        except Exception as e:
            logger.warning(f"[LOGIN] No se pudo registrar audit fallido: {e}")
        messages.error(request, "Usuario no autorizado o credenciales incorrectas.")
        return render(request, "login/login.html", {"form": form})

    if response.status_code != 200:
        messages.error(request, "Error inesperado. Intente más tarde.")
        return render(request, "login/login.html", {"form": form})

    data         = response.json()
    access_token = data["access"]
    user_data    = data["user"]

    clear_failed_attempts(username, request)
    logger.info("Login exitoso via API: '%s'", username)

    # Audit login exitoso
    try:
        nombre_completo = f"{user_data.get('first_name', '')} {user_data.get('last_name', '')}".strip() or username
        _post_audit({
            "username":    username,
            "usuario_id":  user_data.get("id"),
            "accion":      "LOGIN_OK",
            "modulo":      "SISTEMA",
            "severidad":   "INFO",
            "descripcion": f"Inicio de sesión exitoso - {nombre_completo}",
            "ip_address":  request.META.get("REMOTE_ADDR", ""),
            "user_agent":  request.META.get("HTTP_USER_AGENT", "")[:500],
            "url":         request.path,
            "metodo_http": request.method,
        })
    except Exception as e:
        logger.warning(f"[LOGIN] No se pudo registrar audit exitoso: {e}")

    next_url = request.GET.get("next", "")
    if not next_url or not next_url.startswith("/") or next_url.startswith("//"):
        next_url = "/dashboard/"

    resp = redirect(next_url)

    # Guardar el JWT en cookie httponly
    resp.set_cookie(
        "hz_token",
        access_token,
        max_age=8 * 60 * 60,
        httponly=True,
        samesite="Lax",
        secure=settings.SESSION_COOKIE_SECURE,
    )

    return resp


@never_cache
def logout_view(request):
    # Audit logout
    if request.user.is_authenticated:
        try:
            _post_audit({
                "username":    request.user.username,
                "usuario_id":  request.user.id,
                "accion":      "LOGOUT",
                "modulo":      "SISTEMA",
                "severidad":   "INFO",
                "descripcion": f"Cierre de sesión - {request.user.get_full_name() or request.user.username}",
                "ip_address":  request.META.get("REMOTE_ADDR", ""),
                "user_agent":  request.META.get("HTTP_USER_AGENT", "")[:500],
                "url":         request.path,
                "metodo_http": request.method,
            })
        except Exception as e:
            logger.warning(f"[LOGOUT] No se pudo registrar audit: {e}")

    logger.info("Usuario cerró sesión.")
    resp = redirect("usuarios:login")
    resp.delete_cookie("hz_token")
    return resp