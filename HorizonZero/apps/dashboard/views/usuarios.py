# apps/dashboard/views/usuarios.py
# ─────────────────────────────────────────────────────────────────
# Contiene todo el CRUD de agentes + vistas de QR.
# ─────────────────────────────────────────────────────────────────
import base64
import io
import zipfile

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from apps.core.decorators import permiso_requerido
from apps.core.audit import audit
from apps.dashboard.services.agentes_service import (
    AgentesService, UNIDAD_ENCUESTA, ENCUESTA_NOMBRE,
    build_encuesta_url, generar_qr_base64,
)


@login_required
@permiso_requerido("dashboard.view_agentes")
def agentes_home(request):
    buscar      = request.GET.get("q", "").strip()
    unidad_id   = request.GET.get("unidad_id") or None
    sucursal_id = request.GET.get("sucursal_id") or None

    agentes_lista = AgentesService.listar(buscar, unidad_id, sucursal_id)
    sucursales    = AgentesService.sucursales()
    unidades      = AgentesService.unidades()
    kpis          = AgentesService.kpis()
    conteo        = AgentesService.conteo_por_unidad()

    paginator = Paginator(agentes_lista, 20)
    page_obj  = paginator.get_page(request.GET.get("page", 1))

    return render(request, "dashboard/usuarios/agentes_home.html", {
        "agentes":         page_obj,
        "sucursales":      sucursales,
        "unidades":        unidades,
        "kpis":            kpis,
        "conteo":          conteo,
        "buscar":          buscar,
        "filtro_unidad":   unidad_id,
        "filtro_sucursal": sucursal_id,
        "paginator":       paginator,
        "page_obj":        page_obj,
    })


@login_required
@permiso_requerido("dashboard.add_agente")
def agente_crear(request):
    sucursales = AgentesService.sucursales()
    unidades   = AgentesService.unidades()

    if request.method == "POST":
        nombre      = request.POST.get("nombre", "").strip()
        sucursal_id = request.POST.get("sucursal_id")
        unidad_id   = request.POST.get("unidad_id")
        login_ad    = request.POST.get("login_ad", "").strip().lower()

        if not nombre or not sucursal_id or not unidad_id:
            messages.error(request, "Todos los campos son requeridos.")
        else:
            if login_ad and login_ad != "pendiente":
                existente = AgentesService.buscar_por_login_ad(login_ad)
                if existente:
                    return render(request, "dashboard/usuarios/agente_form.html", {
                        "sucursales": sucursales, "unidades": unidades,
                        "modo": "crear", "agente_existente": existente,
                        "login_ad_buscado": login_ad,
                    })
            try:
                nuevo_id = AgentesService.crear(
                    nombre, int(sucursal_id), int(unidad_id), login_ad or "pendiente"
                )
                audit(
                    request, "AGENTE_CREATE", "AGENTES",
                    f"Creó el agente '{nombre}' (login: {login_ad or 'pendiente'})",
                    objeto_id=nuevo_id, objeto_nombre=nombre,
                    datos_nuevos={
                        "nombre": nombre, "unidad_id": unidad_id,
                        "sucursal_id": sucursal_id, "login_ad": login_ad,
                    }
                )
                messages.success(request, f"Agente '{nombre}' creado con ID #{nuevo_id}.")
                return redirect("dashboard:agentes_home")
            except Exception as e:
                messages.error(request, f"Error al crear: {e}")

    return render(request, "dashboard/usuarios/agente_form.html", {
        "sucursales": sucursales, "unidades": unidades, "modo": "crear",
    })


@login_required
@permiso_requerido("dashboard.change_agente")
def agente_editar(request, agente_id):
    agente = AgentesService.obtener(agente_id)
    if not agente:
        messages.error(request, "Agente no encontrado.")
        return redirect("dashboard:agentes_home")

    sucursales = AgentesService.sucursales()
    unidades   = AgentesService.unidades()

    if request.method == "POST":
        nombre      = request.POST.get("nombre", "").strip()
        sucursal_id = request.POST.get("sucursal_id")
        unidad_id   = request.POST.get("unidad_id")
        login_ad    = request.POST.get("login_ad", "").strip().lower()

        if not nombre or not sucursal_id or not unidad_id:
            messages.error(request, "Todos los campos son requeridos.")
        else:
            if login_ad and login_ad != "pendiente":
                existente = AgentesService.buscar_por_login_ad(login_ad)
                if existente and existente["agente_id"] != agente_id:
                    messages.error(
                        request,
                        f"El login '{login_ad}' ya está asignado al agente #{existente['agente_id']}."
                    )
                    return render(request, "dashboard/usuarios/agente_form.html", {
                        "agente": agente, "sucursales": sucursales,
                        "unidades": unidades, "modo": "editar",
                    })
            try:
                AgentesService.actualizar(
                    agente_id, nombre, int(sucursal_id), int(unidad_id), login_ad
                )
                audit(
                    request, "AGENTE_UPDATE", "AGENTES",
                    f"Editó el agente '{agente['nombre']}'",
                    objeto_id=agente_id, objeto_nombre=nombre,
                    datos_anteriores={
                        "nombre": agente["nombre"], "login_ad": agente["login_ad"]
                    },
                    datos_nuevos={"nombre": nombre, "login_ad": login_ad}
                )
                messages.success(request, f"Agente #{agente_id} actualizado.")
                return redirect("dashboard:agentes_home")
            except Exception as e:
                messages.error(request, f"Error al actualizar: {e}")

    return render(request, "dashboard/usuarios/agente_form.html", {
        "agente": agente, "sucursales": sucursales, "unidades": unidades, "modo": "editar",
    })


@login_required
@permiso_requerido("dashboard.delete_agente")
@require_POST
def agente_eliminar(request, agente_id):
    try:
        agente = AgentesService.obtener(agente_id)
        AgentesService.eliminar(agente_id)
        audit(
            request, "AGENTE_DELETE", "AGENTES",
            f"Eliminó el agente '{agente['nombre']}' (soft delete)",
            objeto_id=agente_id, objeto_nombre=agente["nombre"],
            severidad="WARNING"
        )
        messages.success(request, f"Agente #{agente_id} eliminado.")
    except Exception as e:
        messages.error(request, f"Error al eliminar: {e}")
    return redirect("dashboard:agentes_home")


@login_required
@permiso_requerido("dashboard.view_agentes")
def agente_qr(request, agente_id):
    agente = AgentesService.obtener(agente_id)
    if not agente:
        messages.error(request, "Agente no encontrado.")
        return redirect("dashboard:agentes_home")

    qr_data = AgentesService.qr_data(agente_id, agente["unidad_nombre"])
    return render(request, "dashboard/usuarios/agente_qr.html", {
        "agente": agente, "qr_data": qr_data,
    })


@login_required
@permiso_requerido("dashboard.view_agentes")
def agente_qr_download(request, agente_id, encuesta_id):
    agente = AgentesService.obtener(agente_id)
    if not agente:
        return HttpResponse("Not found", status=404)

    unidad_key = (agente["unidad_nombre"] or "").strip().lower()
    encuesta_id = UNIDAD_ENCUESTA.get(unidad_key, 1)

    url     = build_encuesta_url(int(encuesta_id), agente_id)
    qr_b64  = generar_qr_base64(url, size=12)
    png     = base64.b64decode(qr_b64)

    enc_nombre     = ENCUESTA_NOMBRE.get(int(encuesta_id), f"encuesta_{encuesta_id}")
    nombre_archivo = f"QR_{agente['nombre'].replace(' ', '_')}_{enc_nombre.replace(' ', '_')}.png"

    audit(
        request, "EXPORT_QR", "AGENTES",
        f"Descargó QR del agente '{agente['nombre']}' — {enc_nombre}",
        objeto_id=agente_id, objeto_nombre=agente["nombre"]
    )

    response = HttpResponse(png, content_type="image/png")
    response["Content-Disposition"] = f'attachment; filename="{nombre_archivo}"'
    return response


@login_required
@permiso_requerido("dashboard.view_agentes")
def agente_qr_download_zip(request, agente_id):
    agente = AgentesService.obtener(agente_id)
    if not agente:
        return HttpResponse("Not found", status=404)

    qr_data = AgentesService.qr_data(agente_id, agente["unidad_nombre"])
    buffer  = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in qr_data:
            png   = base64.b64decode(item["qr_b64"])
            fname = f"QR_{item['nombre'].replace(' ', '_')}.png"
            zf.writestr(fname, png)

    buffer.seek(0)
    nombre_zip = f"QR_{agente['nombre'].replace(' ', '_')}.zip"

    audit(
        request, "EXPORT_ZIP", "AGENTES",
        f"Descargó ZIP de QR del agente '{agente['nombre']}'",
        objeto_id=agente_id, objeto_nombre=agente["nombre"]
    )

    response = HttpResponse(buffer, content_type="application/zip")
    response["Content-Disposition"] = f'attachment; filename="{nombre_zip}"'
    return response


@login_required
@permiso_requerido("dashboard.view_agentes")
def agentes_inactivos(request):
    buscar      = request.GET.get("q", "").strip()
    unidad_id   = request.GET.get("unidad_id") or None
    sucursal_id = request.GET.get("sucursal_id") or None

    agentes_lista = AgentesService.listar_inactivos(buscar, unidad_id, sucursal_id)
    sucursales    = AgentesService.sucursales()
    unidades      = AgentesService.unidades()

    paginator = Paginator(agentes_lista, 50)
    page_obj  = paginator.get_page(request.GET.get("page", 1))

    return render(request, "dashboard/usuarios/agentes_inactivos.html", {
        "agentes":         page_obj,
        "sucursales":      sucursales,
        "unidades":        unidades,
        "buscar":          buscar,
        "filtro_unidad":   unidad_id,
        "filtro_sucursal": sucursal_id,
        "paginator":       paginator,
        "page_obj":        page_obj,
    })


@login_required
@permiso_requerido("dashboard.delete_agente")
@require_POST
def agente_restaurar(request, agente_id):
    try:
        AgentesService.restaurar(agente_id)
        audit(
            request, "AGENTE_RESTORE", "AGENTES",
            f"Restauró el agente #{agente_id}",
            objeto_id=agente_id,
            severidad="WARNING"
        )
        messages.success(request, f"Agente #{agente_id} restaurado correctamente.")
    except Exception as e:
        messages.error(request, f"Error al restaurar: {e}")
    return redirect("dashboard:agentes_inactivos")