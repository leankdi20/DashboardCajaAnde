# apps/dashboard/views/logs.py
# ─────────────────────────────────────────────────────────────────
# Contiene: logs_home, log_detalle, logs_export_excel, logs_export_pdf
# ─────────────────────────────────────────────────────────────────
import io
import json

from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db.models import Count, Q
from django.db.models.functions import TruncDate
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.contrib import messages

from apps.core.decorators import permiso_requerido
from apps.dashboard.models import AuditLog

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


def _apply_log_filters(qs, params):
    if params.get("usuario"):
        qs = qs.filter(username__icontains=params["usuario"])
    if params.get("accion"):
        qs = qs.filter(accion=params["accion"])
    if params.get("modulo"):
        qs = qs.filter(modulo=params["modulo"])
    if params.get("severidad"):
        qs = qs.filter(severidad=params["severidad"])
    if params.get("ip"):
        qs = qs.filter(ip_address__icontains=params["ip"])
    if params.get("desde"):
        qs = qs.filter(fecha__date__gte=params["desde"])
    if params.get("hasta"):
        qs = qs.filter(fecha__date__lte=params["hasta"])
    return qs


@login_required
@permiso_requerido("dashboard.view_logs")
def logs_home(request):
    qs = _apply_log_filters(AuditLog.objects.all(), request.GET)

    total_eventos     = qs.count()
    logins_fallidos   = qs.filter(accion=AuditLog.Accion.LOGIN_FAIL).count()
    acciones_criticas = qs.filter(severidad=AuditLog.Severidad.CRITICAL).count()
    usuarios_unicos   = qs.values("username").distinct().count()

    actividad = (
        qs.annotate(dia=TruncDate("fecha"))
        .values("dia").annotate(total=Count("id"))
        .order_by("dia")[:30]
    )
    dist_accion = (
        qs.values("accion").annotate(total=Count("id")).order_by("-total")[:8]
    )

    paginator = Paginator(qs, 10)
    page_obj  = paginator.get_page(request.GET.get("page", 1))

    return render(request, "dashboard/logs/logs_home.html", {
        "logs":               page_obj,
        "paginator":          paginator,
        "page_obj":           page_obj,
        "filtro_usuario":     request.GET.get("usuario", ""),
        "filtro_accion":      request.GET.get("accion", ""),
        "filtro_modulo":      request.GET.get("modulo", ""),
        "filtro_severidad":   request.GET.get("severidad", ""),
        "filtro_ip":          request.GET.get("ip", ""),
        "filtro_desde":       request.GET.get("desde", ""),
        "filtro_hasta":       request.GET.get("hasta", ""),
        "total_eventos":      total_eventos,
        "logins_fallidos":    logins_fallidos,
        "acciones_criticas":  acciones_criticas,
        "usuarios_unicos":    usuarios_unicos,
        "grafico_labels":     json.dumps([str(r["dia"]) for r in actividad]),
        "grafico_data":       json.dumps([r["total"] for r in actividad]),
        "dist_labels":        json.dumps([r["accion"] for r in dist_accion]),
        "dist_data":          json.dumps([r["total"] for r in dist_accion]),
        "acciones":           AuditLog.Accion.choices,
        "modulos":            AuditLog.Modulo.choices,
        "severidades":        AuditLog.Severidad.choices,
    })


@login_required
@permiso_requerido("dashboard.view_logs")
def log_detalle(request, log_id):
    try:
        log = AuditLog.objects.get(pk=log_id)
    except AuditLog.DoesNotExist:
        messages.error(request, "Registro no encontrado.")
        return redirect("dashboard:logs_home")

    def _parse(raw):
        if not raw:
            return None
        try:
            return json.loads(raw)
        except Exception:
            return raw

    return render(request, "dashboard/logs/log_detalle.html", {
        "log":              log,
        "datos_anteriores": _parse(log.datos_anteriores),
        "datos_nuevos":     _parse(log.datos_nuevos),
    })


@login_required
@permiso_requerido("dashboard.view_logs")
def logs_export_excel(request):
    qs = _apply_log_filters(AuditLog.objects.all(), request.GET)

    wb = Workbook(); ws = wb.active; ws.title = "Logs de Auditoría"
    AZUL = "003FB7"; GRIS = "F1F5F9"

    headers = ["ID", "Fecha / Hora", "Usuario", "Acción", "Módulo",
               "Severidad", "Descripción", "Objeto ID", "Objeto Nombre",
               "IP Address", "URL", "Método HTTP"]

    hf = PatternFill("solid", fgColor=AZUL)
    hfont = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = hf; c.font = hfont
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    for ri, log in enumerate(qs, 2):
        fecha_str = log.fecha.strftime("%d/%m/%Y %H:%M:%S") if log.fecha else ""
        if log.severidad == "CRITICAL":
            rf = PatternFill("solid", fgColor="FEE2E2")
        elif log.severidad == "WARNING":
            rf = PatternFill("solid", fgColor="FEF9C3")
        else:
            rf = PatternFill("solid", fgColor="FFFFFF") if ri % 2 == 0 else PatternFill("solid", fgColor=GRIS)

        for col, val in enumerate([
            str(log.id), fecha_str, log.username,
            log.get_accion_display(), log.get_modulo_display(), log.get_severidad_display(),
            log.descripcion, log.objeto_id or "", log.objeto_nombre or "",
            str(log.ip_address) if log.ip_address else "", log.url or "", log.metodo_http or "",
        ], 1):
            c = ws.cell(row=ri, column=col, value=val)
            c.fill = rf; c.font = Font(name="Arial", size=9)
            c.alignment = Alignment(vertical="center")
        ws.row_dimensions[ri].height = 18

    for i, w in enumerate([8,18,20,22,16,12,60,10,30,16,50,10], 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"

    # Hoja resumen
    ws2 = wb.create_sheet("Resumen USI")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 20
    ws2.cell(row=1, column=1, value="REPORTE DE AUDITORÍA — HorizonZero").font = Font(bold=True, color=AZUL, size=13)
    ws2.cell(row=2, column=1, value=f"Generado el {dt.now().strftime('%d/%m/%Y %H:%M')}").font = Font(color="64748B", size=9)

    for i, (label, valor) in enumerate([
        ("Total de eventos",         qs.count()),
        ("Logins exitosos",          qs.filter(accion="LOGIN_OK").count()),
        ("Logins fallidos",          qs.filter(accion="LOGIN_FAIL").count()),
        ("Eventos críticos",         qs.filter(severidad="CRITICAL").count()),
        ("Eventos advertencia",      qs.filter(severidad="WARNING").count()),
        ("Usuarios únicos",          qs.values("username").distinct().count()),
        ("Acciones sobre agentes",   qs.filter(modulo="AGENTES").count()),
        ("Exportaciones",            qs.filter(accion__in=["EXPORT_EXCEL","EXPORT_QR","EXPORT_ZIP"]).count()),
    ], 4):
        ws2.cell(row=i, column=1, value=label).font = Font(name="Arial", size=10)
        c = ws2.cell(row=i, column=2, value=valor)
        c.font = Font(bold=True, color=AZUL, size=10)
        c.alignment = Alignment(horizontal="center")

    fstr = dt.now().strftime("%Y%m%d_%H%M")
    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="audit_log_{fstr}.xlsx"'
    wb.save(response); return response


@login_required
@permiso_requerido("dashboard.view_logs")
def logs_export_pdf(request):
    try:
        qs = _apply_log_filters(AuditLog.objects.all(), request.GET)[:500]

        buffer = io.BytesIO()
        doc    = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                                   rightMargin=1*cm, leftMargin=1*cm,
                                   topMargin=1.5*cm, bottomMargin=1*cm)

        styles = getSampleStyleSheet()
        azul   = colors.HexColor("#003FB7")
        gris   = colors.HexColor("#F1F5F9")

        titulo_style = ParagraphStyle("titulo", parent=styles["Heading1"],
                                      textColor=azul, fontSize=14, spaceAfter=4)
        sub_style    = ParagraphStyle("sub", parent=styles["Normal"],
                                      textColor=colors.HexColor("#64748B"), fontSize=8, spaceAfter=12)
        celda_style  = ParagraphStyle("celda", parent=styles["Normal"], fontSize=7)

        elementos = [
            Paragraph("Reporte de Auditoría — HorizonZero · Caja de ANDE", titulo_style),
            Paragraph(
                f"Generado el {dt.now().strftime('%d/%m/%Y %H:%M')} "
                f"por {request.user.get_full_name() or request.user.username} · "
                "Clasificación: USO INTERNO RESTRINGIDO",
                sub_style,
            ),
        ]

        data = [["ID", "Fecha / Hora", "Usuario", "Acción", "Módulo", "Severidad", "Descripción", "IP"]]
        SEV_COLOR = {
            "CRITICAL": colors.HexColor("#FEE2E2"),
            "WARNING":  colors.HexColor("#FEF9C3"),
            "INFO":     colors.white,
        }
        row_colors = [colors.HexColor("#003FB7")]

        for log in qs:
            fecha_str = log.fecha.strftime("%d/%m/%Y\n%H:%M:%S") if log.fecha else ""
            data.append([
                str(log.id), fecha_str, log.username,
                Paragraph(log.get_accion_display(), celda_style),
                log.get_modulo_display(), log.get_severidad_display(),
                Paragraph(log.descripcion[:80], celda_style),
                str(log.ip_address) if log.ip_address else "—",
            ])
            row_colors.append(SEV_COLOR.get(log.severidad, colors.white))

        col_widths = [1.2*cm, 2.5*cm, 3*cm, 3.5*cm, 2.5*cm, 2*cm, 8*cm, 2.5*cm]
        tabla = Table(data, colWidths=col_widths, repeatRows=1)
        tabla.setStyle(TableStyle([
            ("BACKGROUND",   (0,0), (-1,0), azul),
            ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
            ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",     (0,0), (-1,0), 7),
            ("ALIGN",        (0,0), (-1,0), "CENTER"),
            ("BOTTOMPADDING",(0,0), (-1,0), 6),
            ("FONTNAME",     (0,1), (-1,-1), "Helvetica"),
            ("FONTSIZE",     (0,1), (-1,-1), 7),
            ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.white, gris]),
            ("GRID",         (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ("LEFTPADDING",  (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ]))
        for i, color in enumerate(row_colors[1:], 1):
            if color != colors.white:
                tabla.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), color)]))

        elementos.append(tabla)
        doc.build(elementos)
        buffer.seek(0)

        fstr = dt.now().strftime("%Y%m%d_%H%M")
        response = HttpResponse(buffer, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="audit_log_{fstr}.pdf"'
        return response

    except ImportError:
        messages.error(request, "ReportLab no está instalado. Ejecutá: pip install reportlab")
        return redirect("dashboard:logs_home")
