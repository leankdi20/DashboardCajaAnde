from django.core.management.base import BaseCommand
from django.core.mail import EmailMultiAlternatives
from django.conf import settings
from django.utils import timezone
import datetime
from apps.dashboard.models import DetractorNotificado
from apps.dashboard.reports.encuesta_satisfaccion import ReporteEncuestaSatisfaccion


class Command(BaseCommand):
    help = "Revisa detractores nuevos y notifica por correo a la jefatura"

    def handle(self, *args, **options):
        self.stdout.write(f"[{timezone.now():%d/%m/%Y %H:%M}] Revisando detractores...")

        # 1. Obtener IDs ya notificados
        ya_notificados = set(
            DetractorNotificado.objects.values_list("respuesta_id", flat=True)
        )

        # 2. Obtener todos los detractores actuales
        datos = ReporteEncuestaSatisfaccion.obtener_datos(
            {"clasificacion": "detractor"}
        )

        # 3. Agrupar por respuesta_id (una fila por encuesta)
        detractores = {}
        for fila in datos:
            rid = fila["respuesta_id"]
            if rid not in detractores:
                detractores[rid] = fila

        # 4. Filtrar solo los nuevos
        nuevos = [d for rid, d in detractores.items() if rid not in ya_notificados]

        if not nuevos:
            self.stdout.write("  No hay detractores nuevos.")
            return

        self.stdout.write(f"  Encontrados {len(nuevos)} detractores nuevos.")

        destino = settings.DETRACTOR_EMAIL_DESTINO
        enviados = 0
        errores  = 0

        for det in nuevos:
            rid = det["respuesta_id"]
            detalle = ReporteEncuestaSatisfaccion.obtener_detalle(rid)
            
            # Convertir fecha — hacer aquí, antes del try/except
            fecha_raw = det.get("Fecha")
            if fecha_raw and isinstance(fecha_raw, datetime.datetime):
                if timezone.is_naive(fecha_raw):
                    fecha_raw = timezone.make_aware(fecha_raw)

            try:
                self._enviar_correo(det, detalle, destino)
                DetractorNotificado.objects.create(
                    respuesta_id   = rid,
                    encuesta_id    = det.get("encuesta_id"),
                    agente         = det.get("Agente", ""),
                    sucursal       = det.get("Sucursal", ""),
                    promedio       = det.get("promedio_encuesta"),
                    fecha_encuesta = fecha_raw,    # ← usar fecha_raw, no det.get("Fecha")
                    correo_enviado = True,
                )
                enviados += 1
                self.stdout.write(f"  ✓ Notificado: respuesta_id={rid} — {det.get('Agente')}")

            except Exception as e:
                DetractorNotificado.objects.create(
                    respuesta_id   = rid,
                    encuesta_id    = det.get("encuesta_id"),
                    agente         = det.get("Agente", ""),
                    sucursal       = det.get("Sucursal", ""),
                    promedio       = det.get("promedio_encuesta"),
                    fecha_encuesta = fecha_raw,    # ← ya correcto
                    correo_enviado = False,
                    error_envio    = str(e)[:500],
                )
                errores += 1
                self.stdout.write(
                    self.style.ERROR(f"  ✗ Error en respuesta_id={rid}: {e}")
                )

        self.stdout.write(
            self.style.SUCCESS(
                f"  Completado: {enviados} enviados, {errores} errores."
            )
        )

    def _enviar_correo(self, det, detalle, destino):
        agente   = det.get("Agente", "Sin nombre")
        sucursal = det.get("Sucursal", "")
        unidad   = det.get("Unidad", "")
        fecha    = det.get("Fecha", "")
        promedio = det.get("promedio_encuesta", 0)
        nombre_accionista = det.get("Nombre", "No registrado")
        cedula   = det.get("Cedula", "No registrada")
        correo   = det.get("Correo", "No registrado")
        gestion  = det.get("Gestion", "No registrada")

        asunto = f"⚠️ Alerta Detractor — {agente} | {sucursal} | {fecha}"

        # Construir tabla de preguntas/respuestas
        filas_detalle = ""
        for fila in detalle:
            filas_detalle += f"""
                <tr>
                    <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;
                               color:#475569;font-size:13px;">{fila.get('Pregunta','')}</td>
                    <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;
                               font-weight:600;color:#1e293b;font-size:13px;">{fila.get('Respuesta','')}</td>
                </tr>
            """

        html = f"""
        <!DOCTYPE html>
        <html>
        <body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f8fafc;">
          <div style="max-width:640px;margin:32px auto;background:#ffffff;
                      border-radius:12px;overflow:hidden;
                      box-shadow:0 4px 24px rgba(0,0,0,0.08);">

            <!-- Header -->
            <div style="background:#003FB7;padding:24px 32px;">
              <p style="margin:0;color:#93c5fd;font-size:12px;font-weight:600;
                        text-transform:uppercase;letter-spacing:1px;">
                Alerta Automática — HorizonZero
              </p>
              <h1 style="margin:8px 0 0;color:#ffffff;font-size:22px;">
                ⚠️ Nuevo Detractor Registrado
              </h1>
            </div>

            <!-- Promedio badge -->
            <div style="background:#FEF2F2;border-left:4px solid #DC2626;
                        padding:16px 32px;display:flex;align-items:center;gap:16px;">
              <div style="background:#DC2626;color:#fff;border-radius:50%;
                          width:48px;height:48px;display:flex;align-items:center;
                          justify-content:center;font-size:18px;font-weight:700;
                          flex-shrink:0;">{round(promedio, 1)}</div>
              <div>
                <p style="margin:0;font-size:13px;color:#991B1B;font-weight:600;">
                  Promedio de satisfacción: {round(promedio, 2)} / 5
                </p>
                <p style="margin:4px 0 0;font-size:12px;color:#B91C1C;">
                  Clasificación: DETRACTOR (promedio ≤ 2)
                </p>
              </div>
            </div>

            <div style="padding:24px 32px;">

              <!-- Datos del agente -->
              <h2 style="margin:0 0 12px;font-size:15px;color:#003FB7;
                         border-bottom:2px solid #E8EEFF;padding-bottom:8px;">
                Datos del Agente
              </h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;width:140px;">Agente</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;font-weight:600;">{agente}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;">Sucursal</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{sucursal}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;">Unidad</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{unidad}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;">Fecha encuesta</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{fecha}</td>
                </tr>
              </table>

              <!-- Datos del accionista -->
              <h2 style="margin:0 0 12px;font-size:15px;color:#003FB7;
                         border-bottom:2px solid #E8EEFF;padding-bottom:8px;">
                Datos del Accionista
              </h2>
              <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;width:140px;">Nombre</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{nombre_accionista}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;">Cédula</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{cedula}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;">Correo</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{correo}</td>
                </tr>
                <tr>
                  <td style="padding:6px 0;color:#64748B;font-size:13px;">Gestión</td>
                  <td style="padding:6px 0;color:#1e293b;font-size:13px;">{gestion}</td>
                </tr>
              </table>

              <!-- Respuestas de la encuesta -->
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
                <tbody>
                  {filas_detalle}
                </tbody>
              </table>

            </div>

            <!-- Footer -->
            <div style="background:#F1F5F9;padding:16px 32px;
                        border-top:1px solid #E2E8F0;">
              <p style="margin:0;font-size:11px;color:#94A3B8;">
                Este correo fue generado automáticamente por HorizonZero — Caja de ANDE.<br>
                No responder a este mensaje.
              </p>
            </div>

          </div>
        </body>
        </html>
        """

        texto_plano = (
            f"ALERTA DETRACTOR — HorizonZero\n"
            f"Agente: {agente} | Sucursal: {sucursal}\n"
            f"Promedio: {round(promedio, 2)} / 5\n"
            f"Accionista: {nombre_accionista} | Cédula: {cedula}\n"
            f"Fecha: {fecha}\n"
        )

        msg = EmailMultiAlternatives(
            subject=asunto,
            body=texto_plano,
            from_email=settings.DEFAULT_FROM_EMAIL,
            to=[destino],
        )
        msg.attach_alternative(html, "text/html")
        msg.send()