# apps/dashboard/views/system_usuarios.py
import logging
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect

from apps.core.decorators import permiso_requerido
from apps.core.audit import audit
from apps.dashboard.services.api_client import APIClient

logger = logging.getLogger(__name__)


@login_required
@permiso_requerido("dashboard.view_usuarios")
def usuarios_home(request):
    buscar = request.GET.get("q", "").strip()
    try:
        usuarios = APIClient.get(
            "admin/usuarios/",
            params={"q": buscar} if buscar else {},
            request=request,
        )
        grupos = APIClient.get("admin/grupos/", request=request)
    except Exception as e:
        messages.error(request, f"Error al obtener usuarios: {e}")
        usuarios = []
        grupos   = []

    return render(request, "dashboard/system_usuarios/usuarios_home.html", {
        "usuarios": usuarios,
        "grupos":   grupos,
        "buscar":   buscar,
    })


@login_required
@permiso_requerido("dashboard.view_usuarios")
def usuario_crear(request):
    try:
        grupos = APIClient.get("admin/grupos/", request=request)
    except Exception:
        grupos = []

    if request.method == "POST":
        username   = request.POST.get("username", "").strip()
        first_name = request.POST.get("first_name", "").strip()
        last_name  = request.POST.get("last_name", "").strip()
        email      = request.POST.get("email", "").strip()
        is_staff   = request.POST.get("is_staff") == "on"
        grupos_ids = request.POST.getlist("grupos")
        unidad     = request.POST.get("unidad", "").strip()
        sucursal   = request.POST.get("sucursal", "").strip()

        if not username:
            messages.error(request, "El nombre de usuario es requerido.")
            return render(request, "dashboard/system_usuarios/usuario_form.html", {
                "grupos": grupos, "modo": "crear",
            })

        try:
            nuevo = APIClient.post("admin/usuarios/crear/", data={
                "username":   username,
                "first_name": first_name,
                "last_name":  last_name,
                "email":      email,
                "password":   "",
                "is_staff":   is_staff,
                "grupos":     [int(g) for g in grupos_ids if g],
                "unidad":     unidad,
                "sucursal":   sucursal,
            }, request=request)

            audit(
                request, "CREATE", "USUARIOS",
                f"Creó el usuario '{username}'",
                objeto_id=nuevo.get("id"),
                objeto_nombre=username,
                datos_nuevos={"username": username, "email": email},
            )
            
            messages.success(request, f"Usuario '{username}' creado correctamente.")
            return redirect("dashboard:usuarios_home")
        except Exception as e:
            print(f">>> ERROR CREAR USUARIO: {e}")
            messages.error(request, f"Error al crear usuario: {e}")

    return render(request, "dashboard/system_usuarios/usuario_form.html", {
        "grupos": grupos, "modo": "crear",
    })


@login_required
@permiso_requerido("dashboard.view_usuarios")
def usuario_editar(request, user_id):
    try:
        usuario = APIClient.get(f"admin/usuarios/{user_id}/", request=request)
        grupos  = APIClient.get("admin/grupos/", request=request)
    except Exception as e:
        messages.error(request, f"Error: {e}")
        return redirect("dashboard:usuarios_home")

    if request.method == "POST":
        first_name = request.POST.get("first_name", "").strip()
        last_name  = request.POST.get("last_name", "").strip()
        email      = request.POST.get("email", "").strip()
        is_active  = request.POST.get("is_active") == "on"
        is_staff   = request.POST.get("is_staff") == "on"
        grupos_ids = request.POST.getlist("grupos")
        unidad     = request.POST.get("unidad", "").strip()
        sucursal   = request.POST.get("sucursal", "").strip()

        try:
            APIClient.put(f"admin/usuarios/{user_id}/editar/", data={
                "first_name": first_name,
                "last_name":  last_name,
                "email":      email,
                "is_active":  is_active,
                "is_staff":   is_staff,
                "grupos":     [int(g) for g in grupos_ids if g],
                "unidad":     unidad,
                "sucursal":   sucursal,
            }, request=request)

            audit(
                request, "UPDATE", "USUARIOS",
                f"Editó el usuario '{usuario['username']}'",
                objeto_id=user_id,
                objeto_nombre=usuario["username"],
            )
            messages.success(request, "Usuario actualizado correctamente.")
            return redirect("dashboard:usuarios_home")
        except Exception as e:
            messages.error(request, f"Error al actualizar: {e}")

    return render(request, "dashboard/system_usuarios/usuario_form.html", {
        "usuario": usuario,
        "grupos":  grupos,
        "modo":    "editar",
    })



@login_required
@permiso_requerido("dashboard.view_usuarios")
def usuario_desactivar(request, user_id):
    try:
        usuario = APIClient.get(f"admin/usuarios/{user_id}/", request=request)
        APIClient.delete(f"admin/usuarios/{user_id}/eliminar/", request=request)
        audit(
            request, "DELETE", "USUARIOS",
            f"Desactivó el usuario '{usuario['username']}'",
            objeto_id=user_id,
            objeto_nombre=usuario["username"],
            severidad="WARNING",
        )
        messages.success(request, f"Usuario '{usuario['username']}' desactivado.")
    except Exception as e:
        messages.error(request, f"Error: {e}")
    return redirect("dashboard:usuarios_home")


@login_required
@permiso_requerido("dashboard.view_usuarios")
def usuario_permisos(request, user_id):
    try:
        usuario  = APIClient.get(f"admin/usuarios/{user_id}/", request=request)
        todos    = APIClient.get("admin/permisos/", request=request)
        actuales = APIClient.get(f"admin/usuarios/{user_id}/permisos/", request=request)
    except Exception as e:
        messages.error(request, f"Error: {e}")
        return redirect("dashboard:usuarios_home")

    if request.method == "POST":
        permisos_ids = [int(p) for p in request.POST.getlist("permisos") if p]
        try:
            APIClient.put(
                f"admin/usuarios/{user_id}/permisos/",
                data={"permisos": permisos_ids},
                request=request,
            )
            audit(
                request, "UPDATE", "USUARIOS",
                f"Actualizó permisos del usuario '{usuario['username']}'",
                objeto_id=user_id,
                objeto_nombre=usuario["username"],
                severidad="WARNING",
            )
            messages.success(request, "Permisos actualizados correctamente.")
            return redirect("dashboard:usuarios_home")
        except Exception as e:
            messages.error(request, f"Error: {e}")

    # Agrupar permisos por categoría para mostrar en el template
    permisos_por_categoria = {}
    for p in todos:
        categoria = p["name"].split(":")[0].strip() if ":" in p["name"] else "General"
        if categoria not in permisos_por_categoria:
            permisos_por_categoria[categoria] = []
        permisos_por_categoria[categoria].append(p)

    permisos_directos_ids = {p["id"] for p in actuales.get("permisos_directos", [])}
    permisos_grupos_ids   = {p["id"] for p in actuales.get("permisos_grupos", [])}

    return render(request, "dashboard/system_usuarios/usuario_permisos.html", {
        "usuario":              usuario,
        "permisos_por_categoria": permisos_por_categoria,
        "permisos_directos_ids":  permisos_directos_ids,
        "permisos_grupos_ids":    permisos_grupos_ids,
    })