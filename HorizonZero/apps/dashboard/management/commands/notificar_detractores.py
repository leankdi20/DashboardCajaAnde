# ═══════════════════════════════════════════════════════════════════
# apps/dashboard/management/commands/notificar_detractores.py
# ═══════════════════════════════════════════════════════════════════

import datetime
import socket

from django.conf import settings
from django.core.mail import EmailMultiAlternatives
from django.core.management.base import BaseCommand
from django.utils import timezone

from apps.dashboard.models import DetractorNotificado
from apps.dashboard.services.db_service import ReportesDBService


# ── Lógica de valoración ─────────────────────────────────────────

VALOR_SAT_SQL = """
    CASE
        WHEN Pregunta IN (
            '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
            '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
            '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
        ) THEN
            CASE Respuesta
                WHEN 'Muy satisfecho'                THEN 5
                WHEN 'Satisfecho'                    THEN 4
                WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                WHEN 'Poco satisfecho'               THEN 2
                WHEN 'Nada satisfecho'               THEN 1
                WHEN 'Muy fácil'                     THEN 5
                WHEN 'Fácil'                         THEN 4
                WHEN 'Ni fácil ni difícil'           THEN 3
                WHEN 'Difícil'                       THEN 2
                WHEN 'Muy difícil'                   THEN 1
                ELSE NULL
            END
        ELSE NULL
    END
"""

VALOR_OD_SQL = """
    CASE Respuesta
        WHEN 'Excelente' THEN 3
        WHEN 'Regular'   THEN 2
        WHEN 'Malo'      THEN 1
        ELSE NULL
    END
"""

VALOR_WA_SQL = """
    CASE WHEN ISNUMERIC(Respuesta) = 1
         THEN CAST(Respuesta AS FLOAT)
         ELSE NULL
    END
"""

VALOR_WEB_SQL = """
    CASE WHEN ISNUMERIC(Respuesta) = 1
         THEN CAST(Respuesta AS FLOAT)
         ELSE NULL
    END
"""


# ── Definición de grupos ─────────────────────────────────────────
# Columnas disponibles por vista (verificadas con SELECT TOP 1 *):
#
# Satisfaccion:    respuesta_id, Fecha, Hora, Cedula, Nombre, Correo,
#                  Agente, Sucursal, Unidad, Gestion, orden, encuesta_det_id
# OficinaDigital:  respuesta_id, Fecha, Cedula, Nombre, Agente,
#                  orden, encuesta_det_id
# PaginaWeb:       respuesta_id, Fecha, Nombre  (sin Hora, sin Cedula,
#                  sin Agente, sin encuesta_det_id)
# WhatsAppAgente:  respuesta_id, Fecha, Hora, Cedula, Nombre, Agente,
#                  encuesta_det_id
# WhatsAppSinAg:   respuesta_id, Fecha, Hora, Cedula, Nombre,
#                  encuesta_det_id

GRUPOS = [
    {
        "nombre":        "Satisfaccion",
        "vista":         "dbo.vw_reporte_encuestas_satisfaccion",
        "valor_sql":     VALOR_SAT_SQL,
        "escala":        5,
        "umbral":        2,
        "operador":      "<=",
        "where_extra":   "",
        "tiene_agente":  True,
        "tiene_cedula":  True,
        "tiene_gestion": True,
        "tiene_sucursal":True,
        "tiene_unidad":  True,
        "tiene_hora":    True,
        "tiene_detid":   True,
        "settings_key":  "DETRACTOR_EMAIL_SAT",
    },
    {
        "nombre":        "OficinaDigital",
        "vista":         "dbo.vw_reporte_encuestas_satisfaccion_oficina_digital",
        "valor_sql":     VALOR_OD_SQL,
        "escala":        3,
        "umbral":        2,
        "operador":      "<",
        "where_extra":   "AND Respuesta IN ('Excelente','Regular','Malo')",
        "tiene_agente":  True,
        "tiene_cedula":  True,
        "tiene_gestion": False,
        "tiene_sucursal":False,
        "tiene_unidad":  False,
        "tiene_hora":    False,
        "tiene_detid":   True,
        "settings_key":  "DETRACTOR_EMAIL_OD",
    },
    {
        "nombre":        "PaginaWeb",
        "vista":         "dbo.vw_reporte_encuestas_satisfaccion_pagina_web",
        "valor_sql":     VALOR_WEB_SQL,
        "escala":        10,
        "umbral":        3,
        "operador":      "<=",
        "where_extra":   "AND Pregunta IN ('¿Considera que el sitio web es amigable y fácil de navegar?','¿Encontró fácilmente la información que buscaba?')",
        "tiene_agente":  False,
        "tiene_cedula":  False,
        "tiene_gestion": False,
        "tiene_sucursal":False,
        "tiene_unidad":  False,
        "tiene_hora":    False,
        "tiene_detid":   False,  # ← PaginaWeb no tiene encuesta_det_id
        "settings_key":  "DETRACTOR_EMAIL_WEB",
    },
    {
        "nombre":        "WhatsAppAgente",
        "vista":         "dbo.vw_reporte_encuestas_satisfaccion_whatsapp_agente",
        "valor_sql":     VALOR_WA_SQL,
        "escala":        5,
        "umbral":        2,
        "operador":      "<=",
        "where_extra":   "AND Pregunta IN ('¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?','¿La persona que le atendió le brindó respuesta a todas sus consultas?','¿Los tiempos de respuesta fueron adecuados?')",
        "tiene_agente":  True,
        "tiene_cedula":  True,
        "tiene_gestion": False,
        "tiene_sucursal":False,
        "tiene_unidad":  False,
        "tiene_hora":    True,
        "tiene_detid":   True,
        "settings_key":  "DETRACTOR_EMAIL_WA_AGENTE",
    },
    {
        "nombre":        "WhatsAppSinAgente",
        "vista":         "dbo.vw_reporte_encuestas_satisfaccion_WhatsApp_sin_agente",
        "valor_sql":     VALOR_WA_SQL,
        "escala":        5,
        "umbral":        2,
        "operador":      "<=",
        "where_extra":   "AND Pregunta IN ('¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?','¿La persona que le atendió le brindó respuesta a todas sus consultas?','¿Los tiempos de respuesta fueron adecuados?')",
        "tiene_agente":  False,
        "tiene_cedula":  True,
        "tiene_gestion": False,
        "tiene_sucursal":False,
        "tiene_unidad":  False,
        "tiene_hora":    True,
        "tiene_detid":   True,
        "settings_key":  "DETRACTOR_EMAIL_WA_SIN_AGENTE",
    },
]


# ── Queries ──────────────────────────────────────────────────────

def build_query_detractores(grupo: dict, limite: int = None) -> str:
    """
    Construye el query de detractores para una vista específica.
    Usa solo las columnas que existen en cada vista.
    """
    vista       = grupo["vista"]
    valor_sql   = grupo["valor_sql"]
    umbral      = grupo["umbral"]
    operador    = grupo["operador"]
    where_extra = grupo.get("where_extra", "")

    # Columnas SELECT opcionales
    sel_agente   = "Agente,"   if grupo["tiene_agente"]   else ""
    sel_cedula   = "Cedula,"   if grupo["tiene_cedula"]   else ""
    sel_gestion  = "Gestion,"  if grupo["tiene_gestion"]  else ""
    sel_sucursal = "Sucursal," if grupo["tiene_sucursal"] else ""
    sel_unidad   = "Unidad,"   if grupo["tiene_unidad"]   else ""
    sel_hora     = "Hora,"     if grupo["tiene_hora"]     else ""

    # GROUP BY — mismas columnas que SELECT (sin alias v)
    gb_agente   = "Agente,"   if grupo["tiene_agente"]   else ""
    gb_cedula   = "Cedula,"   if grupo["tiene_cedula"]   else ""
    gb_gestion  = "Gestion,"  if grupo["tiene_gestion"]  else ""
    gb_sucursal = "Sucursal," if grupo["tiene_sucursal"] else ""
    gb_unidad   = "Unidad,"   if grupo["tiene_unidad"]   else ""
    gb_hora     = "Hora,"     if grupo["tiene_hora"]     else ""

    top_clause = f"TOP {limite}" if limite else ""

    return f"""
        WITH cal AS (
            SELECT
                respuesta_id,
                Fecha,
                {sel_hora}
                {sel_agente}
                {sel_cedula}
                {sel_gestion}
                {sel_sucursal}
                {sel_unidad}
                Nombre,
                ({valor_sql}) AS v
            FROM {vista}
            WHERE 1=1 {where_extra}
        ),
        prom AS (
            SELECT
                respuesta_id,
                Fecha,
                {sel_hora}
                {sel_agente}
                {sel_cedula}
                {sel_gestion}
                {sel_sucursal}
                {sel_unidad}
                Nombre,
                AVG(CAST(v AS FLOAT)) AS promedio_encuesta
            FROM cal
            WHERE v IS NOT NULL
            GROUP BY
                respuesta_id,
                Fecha,
                {gb_hora}
                {gb_agente}
                {gb_cedula}
                {gb_gestion}
                {gb_sucursal}
                {gb_unidad}
                Nombre
        )
        SELECT {top_clause} *
        FROM prom
        WHERE promedio_encuesta {operador} {umbral}
        ORDER BY Fecha DESC
    """


def build_query_detalle(grupo: dict) -> str:
    """Trae preguntas/respuestas completas de una encuesta."""
    orden = "encuesta_det_id" if grupo["tiene_detid"] else "Pregunta"
    return f"""
        SELECT Pregunta, Respuesta
        FROM {grupo['vista']}
        WHERE respuesta_id = %s
        ORDER BY {orden}
    """


# ════════════════════════════════════════════════════════════════
# COMANDO
# ════════════════════════════════════════════════════════════════

class Command(BaseCommand):
    help = "Notifica detractores de todas las encuestas por correo."

    def add_arguments(self, parser):
        parser.add_argument(
            "--solo-marcar",
            action="store_true",
            help="Marca los detractores sin enviar correo (modo prueba).",
        )
        parser.add_argument(
            "--grupo",
            type=str,
            default=None,
            help="Procesar solo un grupo (ej: Satisfaccion).",
        )
        parser.add_argument(
            "--limite",
            type=int,
            default=None,
            help="Limitar cantidad de detractores por grupo (ej: --limite 1).",
        )

    def handle(self, *args, **options):
        socket.setdefaulttimeout(30)
        solo_marcar  = options["solo_marcar"]
        filtro_grupo = options.get("grupo")
        limite       = options.get("limite")

        self.stdout.write(
            f"[{timezone.now():%d/%m/%Y %H:%M}] Revisando detractores..."
            + (f" (límite: {limite} por grupo)" if limite else "")
        )

        total_enviados = 0
        total_errores  = 0

        grupos_a_procesar = [
            g for g in GRUPOS
            if not filtro_grupo or g["nombre"] == filtro_grupo
        ]

        for grupo in grupos_a_procesar:
            enviados, errores = self._procesar_grupo(grupo, solo_marcar, limite)
            total_enviados += enviados
            total_errores  += errores

        self.stdout.write(self.style.SUCCESS(
            f"\n  Total: {total_enviados} enviados, {total_errores} errores."
        ))

    # ── Procesar un grupo ────────────────────────────────────────
    def _procesar_grupo(self, grupo: dict, solo_marcar: bool,
                        limite: int = None) -> tuple:
        nombre = grupo["nombre"]
        self.stdout.write(f"\n  [{nombre}] Procesando...")

        destinos = getattr(settings, grupo["settings_key"], [])
        if not destinos:
            self.stdout.write(
                self.style.WARNING(f"  [{nombre}] Sin correos configurados — omitido.")
            )
            return 0, 0

        # IDs ya notificados en este grupo
        ya_notificados = set(
            DetractorNotificado.objects.filter(grupo=nombre)
            .values_list("respuesta_id", flat=True)
        )

        # Obtener detractores
        try:
            sql  = build_query_detractores(grupo)
            datos = ReportesDBService.ejecutar_query(sql)
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"  [{nombre}] Error en query: {e}"))
            return 0, 1

        # Deduplicar por respuesta_id y filtrar ya notificados
        detractores = {}
        for fila in datos:
            rid = fila["respuesta_id"]
            if rid not in detractores and rid not in ya_notificados:
                detractores[rid] = fila

        nuevos = list(detractores.values())

        # Aplicar límite si se pasó --limite
        if limite and len(nuevos) > limite:
            self.stdout.write(
                f"  [{nombre}] {len(nuevos)} disponibles — procesando solo {limite}."
            )
            nuevos = nuevos[:limite]
        elif not nuevos:
            self.stdout.write(f"  [{nombre}] Sin detractores nuevos.")
            return 0, 0
        else:
            self.stdout.write(f"  [{nombre}] {len(nuevos)} detractores nuevos.")

        enviados = 0
        errores  = 0

        for det in nuevos:
            rid = det["respuesta_id"]

            # Fecha timezone-aware
            fecha_raw = det.get("Fecha")
            if fecha_raw and isinstance(fecha_raw, datetime.datetime):
                if timezone.is_naive(fecha_raw):
                    fecha_raw = timezone.make_aware(fecha_raw)

            # Detalle de la encuesta
            try:
                detalle_sql = build_query_detalle(grupo)
                detalle = ReportesDBService.ejecutar_query(detalle_sql, [rid])
            except Exception:
                detalle = []

            if solo_marcar:
                obj, created = DetractorNotificado.objects.get_or_create(
                    respuesta_id=rid,
                    grupo=nombre,
                    defaults={
                        "encuesta_id": det.get("encuesta_id"),
                        "agente": det.get("Agente", "") or "",
                        "sucursal": det.get("Sucursal", "") or "",
                        "promedio": det.get("promedio_encuesta"),
                        "fecha_encuesta": fecha_raw,
                        "correo_enviado": False,
                        "error_envio": "solo-marcar",
                    }
                )

                if created:
                    enviados += 1
                    self.stdout.write(f"  [{nombre}] OK marcado respuesta_id={rid}")
                else:
                    self.stdout.write(f"  [{nombre}] Ya existia respuesta_id={rid}")

                continue

            try:
                self._enviar_correo(det, detalle, destinos, grupo)

                obj, created = DetractorNotificado.objects.get_or_create(
                    respuesta_id=rid,
                    grupo=nombre,
                    defaults={
                        "encuesta_id": det.get("encuesta_id"),
                        "agente": det.get("Agente", "") or "",
                        "sucursal": det.get("Sucursal", "") or "",
                        "promedio": det.get("promedio_encuesta"),
                        "fecha_encuesta": fecha_raw,
                        "correo_enviado": True,
                    }
                )

                if created:
                    enviados += 1
                    self.stdout.write(f"  [{nombre}] OK respuesta_id={rid}")
                else:
                    self.stdout.write(f"  [{nombre}] Ya existia respuesta_id={rid}")

            except Exception as e:
                obj, created = DetractorNotificado.objects.get_or_create(
                    respuesta_id=rid,
                    grupo=nombre,
                    defaults={
                        "encuesta_id": det.get("encuesta_id"),
                        "agente": det.get("Agente", "") or "",
                        "sucursal": det.get("Sucursal", "") or "",
                        "promedio": det.get("promedio_encuesta"),
                        "fecha_encuesta": fecha_raw,
                        "correo_enviado": False,
                        "error_envio": str(e)[:500],
                    }
                )

                errores += 1
                self.stdout.write(
                    self.style.ERROR(f"  [{nombre}] ERROR respuesta_id={rid}: {e}")
                )

        return enviados, errores

    # ── Construir y enviar correo ────────────────────────────────
    def _enviar_correo(self, det: dict, detalle: list,
                       destinos: list, grupo: dict):
        nombre_grupo = grupo["nombre"]
        escala       = grupo["escala"]
        promedio     = det.get("promedio_encuesta", 0) or 0
        fecha        = det.get("Fecha", "")

        agente   = det.get("Agente", "—")   if grupo["tiene_agente"]   else "—"
        sucursal = det.get("Sucursal", "—") if grupo["tiene_sucursal"] else "—"
        unidad   = det.get("Unidad", "—")   if grupo["tiene_unidad"]   else "—"
        gestion  = det.get("Gestion", "—")  if grupo["tiene_gestion"]  else "—"
        nombre_accionista = det.get("Nombre", "No registrado") or "No registrado"
        cedula   = det.get("Cedula", "—")   if grupo["tiene_cedula"]   else "—"

        asunto = f"⚠️ Detractor [{nombre_grupo}] — {fecha}"
        if agente != "—":
            asunto += f" — {agente}"

        # Filas de respuestas
        filas_detalle = ""
        for fila in detalle:
            pregunta  = fila.get("Pregunta", "")
            respuesta = fila.get("Respuesta", "")
            if "Recomienda" in pregunta or "brindó respuesta" in pregunta:
                respuesta = "Sí" if str(respuesta) == "1" else (
                    "No" if str(respuesta) == "0" else respuesta
                )
            filas_detalle += f"""
                <tr>
                    <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;
                               color:#475569;font-size:13px;">{pregunta}</td>
                    <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;
                               font-weight:600;color:#1e293b;font-size:13px;">{respuesta}</td>
                </tr>"""

        # Sección agente
        filas_agente = ""
        if grupo["tiene_agente"]:
            filas_agente += f'<tr><td style="padding:6px 0;color:#64748B;font-size:13px;width:140px;">Agente</td><td style="padding:6px 0;color:#1e293b;font-size:13px;font-weight:600;">{agente}</td></tr>'
        if grupo["tiene_unidad"]:
            filas_agente += f'<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Unidad</td><td style="padding:6px 0;color:#1e293b;font-size:13px;">{unidad}</td></tr>'
        if grupo["tiene_sucursal"]:
            filas_agente += f'<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Sucursal</td><td style="padding:6px 0;color:#1e293b;font-size:13px;">{sucursal}</td></tr>'
        if grupo["tiene_gestion"]:
            filas_agente += f'<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Gestión</td><td style="padding:6px 0;color:#1e293b;font-size:13px;">{gestion}</td></tr>'

        seccion_agente = ""
        if filas_agente:
            seccion_agente = f"""
              <h2 style="margin:0 0 12px;font-size:15px;color:#003FB7;
                         border-bottom:2px solid #E8EEFF;padding-bottom:8px;">
                Datos del Agente
              </h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
                {filas_agente}
              </table>"""

        promedio_display = f"{round(promedio, 2)} / {escala}"

        html = f"""<!DOCTYPE html>
        <html>
        <body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f8fafc;">
          <div style="max-width:640px;margin:32px auto;background:#ffffff;
                      border-radius:12px;overflow:hidden;
                      box-shadow:0 4px 24px rgba(0,0,0,0.08);">
            <div style="background:#003FB7;padding:24px 32px;">
              <p style="margin:0;color:#93c5fd;font-size:12px;font-weight:600;
                        text-transform:uppercase;letter-spacing:1px;">
                Alerta Automática — HorizonZero · {nombre_grupo}
              </p>
              <h1 style="margin:8px 0 0;color:#ffffff;font-size:22px;">
                ⚠️ Nuevo Detractor Registrado
              </h1>
            </div>
            <div style="background:#FEF2F2;border-left:4px solid #DC2626;
                        padding:16px 32px;">
              <p style="margin:0;font-size:13px;color:#991B1B;font-weight:600;">
                Promedio: {promedio_display}
              </p>
              <p style="margin:4px 0 0;font-size:12px;color:#B91C1C;">
                Clasificación: DETRACTOR — Encuesta: {nombre_grupo}
              </p>
            </div>
            <div style="padding:24px 32px;">
              {seccion_agente}
              <h2 style="margin:0 0 12px;font-size:15px;color:#003FB7;
                         border-bottom:2px solid #E8EEFF;padding-bottom:8px;">
                Datos del Accionista
              </h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
                <tr><td style="padding:6px 0;color:#64748B;font-size:13px;width:140px;">Nombre</td>
                    <td style="padding:6px 0;color:#1e293b;font-size:13px;">{nombre_accionista}</td></tr>
                <tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Cédula</td>
                    <td style="padding:6px 0;color:#1e293b;font-size:13px;">{cedula}</td></tr>
                <tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Fecha</td>
                    <td style="padding:6px 0;color:#1e293b;font-size:13px;">{fecha}</td></tr>
              </table>
              <h2 style="margin:0 0 12px;font-size:15px;color:#003FB7;
                         border-bottom:2px solid #E8EEFF;padding-bottom:8px;">
                Respuestas de la Encuesta
              </h2>
              <table style="width:100%;border-collapse:collapse;
                            background:#F8FAFC;border-radius:8px;overflow:hidden;">
                <thead>
                  <tr style="background:#E8EEFF;">
                    <th style="padding:10px 12px;text-align:left;font-size:12px;
                               color:#003FB7;font-weight:600;">Pregunta</th>
                    <th style="padding:10px 12px;text-align:left;font-size:12px;
                               color:#003FB7;font-weight:600;">Respuesta</th>
                  </tr>
                </thead>
                <tbody>{filas_detalle}</tbody>
              </table>
            </div>
            <div style="background:#F1F5F9;padding:16px 32px;border-top:1px solid #E2E8F0;">
              <p style="margin:0;font-size:11px;color:#94A3B8;">
                Generado automáticamente por HorizonZero — Caja de ANDE.<br>
                No responder a este mensaje.
              </p>
            </div>
          </div>
        </body>
        </html>"""

        texto_plano = (
            f"ALERTA DETRACTOR [{nombre_grupo}] — HorizonZero\n"
            f"Agente: {agente} | Sucursal: {sucursal}\n"
            f"Promedio: {promedio_display}\n"
            f"Accionista: {nombre_accionista} | Cédula: {cedula}\n"
            f"Fecha: {fecha}\n"
        )

        msg = EmailMultiAlternatives(
            subject    = asunto,
            body       = texto_plano,
            from_email = settings.DEFAULT_FROM_EMAIL,
            to         = destinos,
        )
        msg.attach_alternative(html, "text/html")
        msg.send()