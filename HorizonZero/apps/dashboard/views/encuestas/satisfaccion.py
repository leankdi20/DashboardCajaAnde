# apps/dashboard/views/encuestas/satisfaccion.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, JsonResponse

from apps.core.decorators import permiso_requerido
from apps.dashboard.reports.encuesta_satisfaccion import ReporteEncuestaSatisfaccion
from apps.dashboard.services.db_service import ReportesDBService
from apps.dashboard.tables import EncuestaSatisfaccionTable
from apps.dashboard.views._base import build_timeline_heatmap
from django.shortcuts import render


CAMPOS_ENCUESTA_SAT = {
    "agente":   "Agente",
    "unidad":   "Unidad",
    "sucursal": "Sucursal",
    "gestion":  "Gestion",
    "nombre":   "Nombre",
}


@login_required
def encuesta_satisfaccion_buscar(request, campo):
    if campo not in CAMPOS_ENCUESTA_SAT:
        return JsonResponse({"results": []})

    col_db = CAMPOS_ENCUESTA_SAT[campo]
    term   = request.GET.get("term", "").strip()
    if not term:
        return JsonResponse({"results": []})

    sql = f"""
        SELECT DISTINCT TOP 30 {col_db}
        FROM dbo.vw_reporte_encuestas_satisfaccion
        WHERE {col_db} IS NOT NULL
          AND {col_db} LIKE %s
        ORDER BY {col_db}
    """
    try:
        filas = ReportesDBService.ejecutar_query(sql, [f"%{term}%"])
        data  = [
            {"id": (r[col_db] or "").strip(), "text": (r[col_db] or "").strip()}
            for r in filas if r.get(col_db) and r[col_db].strip()
        ]
    except Exception:
        data = []

    return JsonResponse({"results": data, "pagination": {"more": False}})


@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def encuesta_satisfaccion_kpis(request):
    kpi_sucursales   = request.GET.getlist("kpi_sucursal")
    kpi_fecha_inicio = request.GET.get("kpi_fecha_inicio", "").strip() or None
    kpi_fecha_fin    = request.GET.get("kpi_fecha_fin", "").strip() or None

    try:
        kpis = ReporteEncuestaSatisfaccion.obtener_kpis_globales(
            sucursales=kpi_sucursales if kpi_sucursales else None,
            fecha_inicio=kpi_fecha_inicio,
            fecha_fin=kpi_fecha_fin,
        )
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)

    return JsonResponse(kpis)


@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def encuesta_satisfaccion(request):
    filtros = {
        "agente":        request.GET.get("agente"),
        "sucursal":      request.GET.get("sucursales"),
        "unidad":        request.GET.get("unidad"),
        "gestion":       request.GET.get("gestion"),
        "nombre":        request.GET.get("nombre"),
        "cedula":        request.GET.get("cedula"),
        "fecha_inicio":  request.GET.get("fecha_inicio"),
        "fecha_fin":     request.GET.get("fecha_fin"),
        "clasificacion": request.GET.get("clasificacion"),
    }
    kpi_sucursales   = request.GET.getlist("kpi_sucursal")
    kpi_fecha_inicio = request.GET.get("kpi_fecha_inicio", "")
    kpi_fecha_fin    = request.GET.get("kpi_fecha_fin", "")

    datos, opciones_sucursal = [], []
    timeline_data, heatmap_anios, heatmap_json = [], [], "{}"

    try:
        datos = ReporteEncuestaSatisfaccion.obtener_datos_agrupados(filtros)
        opciones_sucursal = ReportesDBService.ejecutar_query(
            "SELECT DISTINCT Sucursal FROM dbo.vw_reporte_encuestas_satisfaccion "
            "WHERE Sucursal IS NOT NULL ORDER BY Sucursal"
        )
        raw_timeline = ReporteEncuestaSatisfaccion.obtener_timeline()
        timeline_data, heatmap_anios, heatmap_json = build_timeline_heatmap(raw_timeline)
    except Exception as e:
        print(">>> ERROR:", e)
        messages.error(request, "Error al obtener los datos.")

    table = EncuestaSatisfaccionTable(datos)
    try:
        table.paginate(page=request.GET.get("page", 1), per_page=10)
    except Exception:
        table.paginate(page=1, per_page=10)

    return render(request, "dashboard/encuestas/satisfaccion.html", {
        "table":             table,
        "filtros":           filtros,
        "kpi_sucursales":    kpi_sucursales,
        "kpi_fecha_inicio":  kpi_fecha_inicio,
        "kpi_fecha_fin":     kpi_fecha_fin,
        "opciones_sucursal": opciones_sucursal,
        "timeline_data":     timeline_data,
        "heatmap_anios":     heatmap_anios,
        "heatmap_json":      heatmap_json,
    })


@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def encuesta_satisfaccion_detalle(request, respuesta_id):
    try:
        filas = ReporteEncuestaSatisfaccion.obtener_detalle(respuesta_id)
    except Exception:
        filas = []
        messages.error(request, "Error al obtener el detalle.")

    for fila in filas:
        pregunta  = (fila.get("Pregunta") or "").lower()
        respuesta = fila.get("Respuesta", "")
        if "recomienda" in pregunta:
            if str(respuesta).strip() == "1":
                fila["Respuesta"] = "Sí"
            elif str(respuesta).strip() in ("0", "2"):
                fila["Respuesta"] = "No"

    encabezado = filas[0] if filas else {}

    respuestas_numericas = []
    for fila in filas:
        try:
            respuestas_numericas.append(float(fila.get("Respuesta", 0)))
        except (ValueError, TypeError):
            pass

    satisfechos      = sum(1 for r in respuestas_numericas if r >= 4)
    satisfaccion_pct = round((satisfechos / len(respuestas_numericas) * 100)) if respuestas_numericas else 0

    stats_agente      = ReporteEncuestaSatisfaccion.obtener_promedio_agente(encabezado.get("Agente", ""))
    promedio_encuesta = ReporteEncuestaSatisfaccion.obtener_promedio_encuesta(respuesta_id)
    promedio_agente_pct   = round((stats_agente["promedio_agente"]  / 5) * 100) if stats_agente["promedio_agente"]  else 0
    promedio_encuesta_pct = round((promedio_encuesta / 5) * 100) if promedio_encuesta else 0

    respuesta_satisfaccion = ""
    for fila in filas:
        if "satisfecho" in (fila.get("Pregunta") or "").lower():
            respuesta_satisfaccion = fila.get("Respuesta", "")
            break

    kpis = {
        "promedio_general":       promedio_agente_pct,
        "total_encuestas":        stats_agente["total_encuestas"],
        "promedio_encuesta":      promedio_encuesta_pct,
        "respuesta_satisfaccion": respuesta_satisfaccion,
        "satisfaccion_pct":       satisfaccion_pct,
        "agente":                 encabezado.get("Agente", ""),
        "sucursal":               encabezado.get("Sucursal", ""),
        "unidad":                 encabezado.get("Unidad", ""),
    }

    return render(request, "dashboard/encuestas/satisfaccion_detalle.html", {
        "encabezado":   encabezado,
        "preguntas":    filas,
        "respuesta_id": respuesta_id,
        "kpis":         kpis,
    })


@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def encuesta_satisfaccion_exportar(request):
    filtros = {
        "agente":        request.GET.get("agente"),
        "sucursal":      request.GET.get("sucursal"),
        "unidad":        request.GET.get("unidad"),
        "gestion":       request.GET.get("gestion"),
        "nombre":        request.GET.get("nombre"),
        "cedula":        request.GET.get("cedula"),
        "fecha_inicio":  request.GET.get("fecha_inicio"),
        "fecha_fin":     request.GET.get("fecha_fin"),
        "clasificacion": request.GET.get("clasificacion"),
    }
    datos = ReporteEncuestaSatisfaccion.obtener_datos(filtros)

    encuestas, preguntas_orden = {}, []
    for fila in datos:
        rid      = fila["respuesta_id"]
        pregunta = fila.get("Pregunta", "")
        respuesta = fila.get("Respuesta", "")

        if rid not in encuestas:
            encuestas[rid] = {
                "respuesta_id": rid,
                "Fecha":    fila.get("Fecha", ""),
                "Hora":     fila.get("Hora", ""),
                "Agente":   fila.get("Agente", ""),
                "Unidad":   fila.get("Unidad", ""),
                "Sucursal": fila.get("Sucursal", ""),
                "Gestion":  fila.get("Gestion", ""),
                "Nombre":   fila.get("Nombre", ""),
                "Cedula":   fila.get("Cedula", ""),
            }
        if pregunta and pregunta not in preguntas_orden:
            preguntas_orden.append(pregunta)
        encuestas[rid][pregunta] = respuesta

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Encuesta Satisfacción"

    cols_fijas = ["ID", "Fecha", "Agente", "Unidad", "Sucursal", "Gestión", "Accionista", "Cédula"]
    keys_fijas = ["respuesta_id", "Fecha", "Agente", "Unidad", "Sucursal", "Gestion", "Nombre", "Cedula"]

    hf = PatternFill("solid", fgColor="003FB7")
    hfont = Font(bold=True, color="FFFFFF")
    for col_idx, col_name in enumerate(cols_fijas, 1):
        c = ws.cell(row=1, column=col_idx, value=col_name)
        c.fill = hf; c.font = hfont
        c.alignment = Alignment(horizontal="center")

    for i, pregunta in enumerate(preguntas_orden, len(cols_fijas) + 1):
        c = ws.cell(row=1, column=i, value=pregunta)
        c.fill = PatternFill("solid", fgColor="FFC900")
        c.font = Font(bold=True, color="1A1000")
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for row_idx, (rid, enc) in enumerate(encuestas.items(), 2):
        for col_idx, key in enumerate(keys_fijas, 1):
            valor = enc.get(key, "")
            ws.cell(row=row_idx, column=col_idx, value=str(valor) if valor else "")
        for i, pregunta in enumerate(preguntas_orden, len(cols_fijas) + 1):
            respuesta = enc.get(pregunta, "")
            if "recomienda" in pregunta.lower():
                if str(respuesta).strip() == "1":
                    respuesta = "Sí"
                elif str(respuesta).strip() in ["0", "2"]:
                    respuesta = "No"
            ws.cell(row=row_idx, column=i, value=str(respuesta) if respuesta else "")

    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    ws.row_dimensions[1].height = 60

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = 'attachment; filename="encuesta_satisfaccion.xlsx"'
    wb.save(response)
    return response


@login_required
@permiso_requerido("dashboard.view_encuesta_satisfaccion")
def encuesta_satisfaccion_detalle_exportar(request, respuesta_id):
    try:
        filas = ReporteEncuestaSatisfaccion.obtener_detalle(respuesta_id)
    except Exception:
        filas = []

    if not filas:
        return HttpResponse("Sin datos", status=404)

    encabezado = filas[0]
    for fila in filas:
        pregunta  = (fila.get("Pregunta") or "").lower()
        respuesta = fila.get("Respuesta", "")
        if "recomienda" in pregunta:
            if str(respuesta).strip() == "1":
                fila["Respuesta"] = "Sí"
            elif str(respuesta).strip() in ["0", "2"]:
                fila["Respuesta"] = "No"

    stats_agente      = ReporteEncuestaSatisfaccion.obtener_promedio_agente(encabezado.get("Agente", ""))
    promedio_encuesta = ReporteEncuestaSatisfaccion.obtener_promedio_encuesta(respuesta_id)
    promedio_agente_pct   = round((stats_agente["promedio_agente"]  / 5) * 100) if stats_agente["promedio_agente"]  else 0
    promedio_encuesta_pct = round((promedio_encuesta / 5) * 100) if promedio_encuesta else 0

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte Encuesta"

    def fill(color): return PatternFill("solid", fgColor=color)
    def border_full(color="CCCCCC"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)
    def font(bold=False, color="1A1A1A", size=10):
        return Font(bold=bold, color=color, size=size, name="Calibri")

    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    der    = Alignment(horizontal="right",  vertical="center", wrap_text=True)

    for col, w in zip("ABCDEFG", [2, 26, 2, 35, 20, 20, 2]):
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 5
    ws.row_dimensions[2].height = 40
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 5

    ws.merge_cells("B2:B3")
    c = ws.cell(row=2, column=2, value="CAJA DE\nANDE")
    c.fill = fill("003FB7"); c.font = font(bold=True, color="FFFFFF", size=10); c.alignment = centro

    ws.merge_cells("D2:E2")
    c = ws.cell(row=2, column=4, value="Reporte de Encuesta de Satisfacción")
    c.font = font(bold=True, color="003FB7", size=15); c.alignment = izq

    ws.merge_cells("D3:E3")
    c = ws.cell(row=3, column=4, value=f"ID Reporte: #{respuesta_id}")
    c.font = font(color="74788A", size=9); c.alignment = izq

    c = ws.cell(row=2, column=6, value="FECHA DE GENERACIÓN")
    c.font = font(bold=True, color="FFC900", size=8); c.alignment = der

    c = ws.cell(row=3, column=6, value=datetime.now().strftime("%d de %B, %Y"))
    c.font = font(bold=True, color="003FB7", size=11); c.alignment = der

    ws.row_dimensions[4].height = 4
    for col in range(2, 7):
        ws.cell(row=4, column=col).fill = fill("003FB7")

    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 40
    ws.row_dimensions[7].height = 5

    kpi_cols   = [4, 5, 6]
    kpi_labels = ["PROMEDIO GENERAL AGENTE", "ENCUESTAS DEL AGENTE", "PROMEDIO ESTA ENCUESTA"]
    kpi_vals   = [f"{promedio_agente_pct}%", str(stats_agente["total_encuestas"]), f"{promedio_encuesta_pct}%"]
    kpi_fills  = ["003FB7", "E8EEFF", "8B1A0A"]
    kpi_fonts  = ["FFFFFF", "003FB7", "FFFFFF"]

    for i, col in enumerate(kpi_cols):
        c = ws.cell(row=5, column=col, value=kpi_labels[i])
        c.fill = fill(kpi_fills[i]); c.font = font(bold=True, color=kpi_fonts[i], size=8)
        c.alignment = centro; c.border = border_full(kpi_fills[i])

        c = ws.cell(row=6, column=col, value=kpi_vals[i])
        c.fill = fill(kpi_fills[i]); c.font = font(bold=True, color=kpi_fonts[i], size=22)
        c.alignment = centro; c.border = border_full(kpi_fills[i])

    panel_items = [
        ("AGENTE RESPONSABLE",   encabezado.get("Agente",   "")),
        ("UNIDAD INSTITUCIONAL", encabezado.get("Unidad",   "")),
        ("SUCURSAL REGIONAL",    encabezado.get("Sucursal", "")),
        ("TIPO DE GESTIÓN",      encabezado.get("Gestion",  "")),
        ("NOMBRE ACCIONISTA",    encabezado.get("Nombre",   "—")),
        ("CÉDULA",               encabezado.get("Cedula",   "—")),
        ("FECHA",                str(encabezado.get("Fecha", ""))),
    ]

    row_enc_tabla = 8
    ws.row_dimensions[row_enc_tabla].height = 18
    ws.merge_cells(start_row=row_enc_tabla, start_column=4, end_row=row_enc_tabla, end_column=5)
    c = ws.cell(row=row_enc_tabla, column=4, value="PREGUNTA")
    c.fill = fill("E8EEFF"); c.font = font(bold=True, color="003FB7", size=9); c.alignment = izq

    c = ws.cell(row=row_enc_tabla, column=6, value="RESPUESTA")
    c.fill = fill("E8EEFF"); c.font = font(bold=True, color="003FB7", size=9); c.alignment = centro

    row_data = 9
    for i in range(max(len(panel_items) * 2, len(filas))):
        ws.row_dimensions[row_data + i].height = 22

    for i, (label, valor) in enumerate(panel_items):
        r_label = row_data + (i * 2)
        r_valor = row_data + (i * 2) + 1
        c = ws.cell(row=r_label, column=2, value=label)
        c.fill = fill("FFF0C0"); c.font = font(bold=True, color="003FB7", size=8); c.alignment = izq
        c = ws.cell(row=r_valor, column=2, value=valor)
        c.fill = fill("FFF8E0"); c.font = font(color="1A1A1A", size=10); c.alignment = izq

    for i, fila in enumerate(filas):
        r = row_data + i
        ws.row_dimensions[r].height = 28
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        c = ws.cell(row=r, column=4, value=fila.get("Pregunta", ""))
        c.fill = fill("FFFFFF") if i % 2 == 0 else fill("F4F7FF")
        c.font = font(color="1A1A1A", size=9); c.alignment = izq; c.border = border_full("E0E0E0")

        respuesta = fila.get("Respuesta", "")
        color_resp = "4CAF50" if respuesta in ["Muy satisfecho", "Muy fácil", "Sí"] else \
                     "F44336" if respuesta in ["No", "Nada satisfecho", "Muy difícil"] else "003FB7"
        c = ws.cell(row=r, column=6, value=respuesta)
        c.fill = fill(color_resp); c.font = font(bold=True, color="FFFFFF", size=9)
        c.alignment = centro; c.border = border_full(color_resp)

    row_pie = row_data + max(len(panel_items) * 2, len(filas)) + 1
    ws.row_dimensions[row_pie].height = 4
    for col in range(2, 7):
        ws.cell(row=row_pie, column=col).fill = fill("FFC900")

    ws.row_dimensions[row_pie + 1].height = 14
    ws.merge_cells(start_row=row_pie + 1, start_column=2, end_row=row_pie + 1, end_column=6)
    c = ws.cell(row=row_pie + 1, column=2, value="© 2026 Caja de Ande · Documento generado automáticamente · Confidencial")
    c.font = font(color="74788A", size=8); c.alignment = centro

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="encuesta_detalle_{respuesta_id}.xlsx"'
    wb.save(response)
    return response