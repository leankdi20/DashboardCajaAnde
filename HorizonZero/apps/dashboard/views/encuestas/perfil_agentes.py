# apps/dashboard/views/encuestas/perfil_agentes.py
import json
from urllib.parse import unquote

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.shortcuts import redirect, render

from apps.core.decorators import permiso_requerido
from apps.dashboard.reports.reporte_perfil_agente import ReportePerfilAgente
from apps.dashboard.reports.reporte_perfil_agente_od import ReportePerfilAgenteOD


PERFIL_CONFIG = {
    "sat": {
        "nombre":      "Satisfacción",
        "color":       "primary",
        "icono":       "star",
        "index_url":   "perfil_agentes_sat_index",
        "data_url":    "perfil_agentes_sat_data",
        "detalle_url": "perfil_agente_sat_detalle",
        "ajax_url":    "perfil_agente_sat_ajax",
        "base_path": "/dashboard/encuestas/agentes/sat/",
        "permiso":     "dashboard.view_encuesta_satisfaccion",
    },
    "od": {
        "nombre":      "Oficina Digital",
        "color":       "blue",
        "icono":       "computer",
        "index_url":   "perfil_agentes_od_index",
        "data_url":    "perfil_agentes_od_data",
        "detalle_url": "perfil_agente_od_detalle",
        "ajax_url":    "perfil_agente_od_ajax",
        "base_path": "/dashboard/encuestas/agentes/od/",
        "permiso":     "dashboard.view_encuesta_satisfaccion_oficina",
    },
}


def _get_service(tipo):
    return ReportePerfilAgente if tipo == "sat" else ReportePerfilAgenteOD


# ── SATISFACCIÓN ──────────────────────────────────────────────────
@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def perfil_agentes_sat_index(request):
    return _perfil_agentes_index(request, "sat")

@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def perfil_agentes_sat_data(request):
    return _perfil_agentes_data(request, "sat")

@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def perfil_agente_sat_detalle(request, agente_nombre):
    return _perfil_agente_detalle(request, agente_nombre, "sat")

@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def perfil_agente_sat_ajax(request, agente_nombre):
    return _perfil_agente_ajax(request, agente_nombre, "sat")


# ── OFICINA DIGITAL ───────────────────────────────────────────────
@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion_oficina")
def perfil_agentes_od_index(request):
    return _perfil_agentes_index(request, "od")

@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion_oficina")
def perfil_agentes_od_data(request):
    return _perfil_agentes_data(request, "od")

@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion_oficina")
def perfil_agente_od_detalle(request, agente_nombre):
    return _perfil_agente_detalle(request, agente_nombre, "od")

@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion_oficina")
def perfil_agente_od_ajax(request, agente_nombre):
    return _perfil_agente_ajax(request, agente_nombre, "od")


# ── LÓGICA COMPARTIDA (privada) ───────────────────────────────────
def _perfil_agentes_index(request, tipo):
    service  = _get_service(tipo)
    config   = PERFIL_CONFIG[tipo]
    opciones = service.obtener_opciones()
    return render(request, "dashboard/encuestas/perfil_agentes_index.html", {
        "opciones":            opciones,
        "filtro_unidad":       request.GET.get("unidad", ""),
        "filtro_sucursal":     request.GET.get("sucursal", ""),
        "filtro_fecha_inicio": request.GET.get("fecha_inicio", ""),
        "filtro_fecha_fin":    request.GET.get("fecha_fin", ""),
        "buscar":              request.GET.get("q", ""),
        "tipo":                tipo,
        "config":              config,
    })


def _perfil_agentes_data(request, tipo):
    service      = _get_service(tipo)
    unidad       = request.GET.get("unidad",       "").strip() or None
    sucursal     = request.GET.get("sucursal",     "").strip() or None
    fecha_inicio = request.GET.get("fecha_inicio", "").strip() or None
    fecha_fin    = request.GET.get("fecha_fin",    "").strip() or None
    buscar       = request.GET.get("q",            "").strip().lower()

    agentes = service.obtener_lista(
        unidad=unidad, sucursal=sucursal,
        fecha_inicio=fecha_inicio, fecha_fin=fecha_fin,
    )
    if buscar:
        agentes = [a for a in agentes if buscar in a["Agente"].lower()]
    return JsonResponse({"agentes": agentes, "total": len(agentes)})


def _perfil_agente_detalle(request, agente_nombre, tipo):
    agente_nombre = unquote(agente_nombre)
    service       = _get_service(tipo)
    config        = PERFIL_CONFIG[tipo]
    fecha_inicio  = request.GET.get("fecha_inicio", "").strip()
    fecha_fin     = request.GET.get("fecha_fin",    "").strip()

    kpis = service.obtener_kpi(
        agente_nombre,
        fecha_inicio=fecha_inicio or None,
        fecha_fin=fecha_fin or None,
    )
    if not kpis:
        messages.error(request, f"No se encontraron datos para '{agente_nombre}'.")
        return redirect(f"dashboard:{config['index_url']}")

    agente_info  = service.obtener_info(agente_nombre)
    ranking_data = service.obtener_ranking(agente_info.get("Unidad", ""), agente_nombre)

    return render(request, "dashboard/encuestas/perfil_agentes_detalle.html", {
        "agente_nombre": agente_nombre,
        "agente_info":   agente_info,
        "kpis":          kpis,
        "ranking_data":  ranking_data,
        "fecha_inicio":  fecha_inicio,
        "fecha_fin":     fecha_fin,
        "tipo":          tipo,
        "config":        config,
    })


def _perfil_agente_ajax(request, agente_nombre, tipo):
    agente_nombre = unquote(agente_nombre)
    service       = _get_service(tipo)
    seccion       = request.GET.get("seccion", "")
    fecha_inicio  = request.GET.get("fecha_inicio", "").strip() or None
    fecha_fin     = request.GET.get("fecha_fin",    "").strip() or None
    escala        = 3 if tipo == "od" else 5

    try:
        if seccion == "tendencia":
            tendencia = service.obtener_tendencia(agente_nombre)
            return JsonResponse({
                "labels":      [t["mes_label"]    for t in tendencia],
                "promedios":   [t["promedio_pct"] for t in tendencia],
                "total_enc":   [t["total_enc"]    for t in tendencia],
                "promotores":  [t["promotores"]   for t in tendencia],
                "pasivos":     [t["pasivos"]       for t in tendencia],
                "detractores": [t["detractores"]  for t in tendencia],
            })

        elif seccion == "gestion":
            dist = service.obtener_dist_gestion(agente_nombre)
            return JsonResponse({
                "labels":    [d["Gestion"] for d in dist],
                "totales":   [d["total"]   for d in dist],
                "promedios": [round((d["promedio"] / escala) * 100) for d in dist],
            })

        elif seccion == "ranking":
            info   = service.obtener_info(agente_nombre)
            rank   = service.obtener_ranking(info.get("Unidad", ""), agente_nombre)
            top10  = rank["ranking"][:10]
            return JsonResponse({
                "nombres":   [r["Agente"]  for r in top10],
                "promedios": [round((r["promedio"] / escala) * 100) for r in top10],
                "es_yo":     [r["Agente"] == agente_nombre for r in top10],
                "posicion":  rank["posicion"],
                "total":     rank["total_agentes"],
            })

        elif seccion == "ultimas":
            ultimas = service.obtener_ultimas(agente_nombre)
            data = [{
                "respuesta_id":  u["respuesta_id"],
                "fecha":         u["Fecha"].strftime("%d/%m/%Y") if u.get("Fecha") else "",
                "nombre":        u.get("Nombre")  or u.get("nombre")  or "",
                "cedula":        u.get("Cedula")  or u.get("cedula")  or "",
                "gestion":       u.get("Gestion") or u.get("gestion") or "",
                "sucursal":      u.get("Sucursal")or u.get("sucursal")or "",
                "promedio":      u.get("promedio", 0),
                "clasificacion": u.get("clasificacion", ""),
            } for u in ultimas]
            return JsonResponse({"ultimas": data})

        return JsonResponse({"error": "Sección no válida"}, status=400)

    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)
