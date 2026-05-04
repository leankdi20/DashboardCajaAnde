# apps/dashboard/views/__init__.py
# ═══════════════════════════════════════════════════════════════════
# Punto de entrada único — reemplaza el antiguo views.py de 6000 líneas.
# urls.py no necesita cambios: sigue importando desde apps.dashboard.views.
# ═══════════════════════════════════════════════════════════════════

from .home import dashboard_home

# ── Encuestas ────────────────────────────────────────────────────
from .encuestas import (
    # Satisfacción
    encuesta_satisfaccion,
    encuesta_satisfaccion_buscar,
    encuesta_satisfaccion_kpis,
    encuesta_satisfaccion_detalle,
    encuesta_satisfaccion_exportar,
    encuesta_satisfaccion_detalle_exportar,
    # Oficina Digital
    encuesta_satisfaccion_oficina,
    encuesta_satisfaccion_detalle_of_dig,
    encuesta_satisfaccion_of_dig_exportar,
    encuesta_satisfaccion_detalle_of_dig_exportar,
    encuesta_oficina_digital_buscar,
    # WhatsApp Agente
    encuesta_whatsApp_agente,
    encuesta_whatsapp_agente_detalle,
    encuesta_whatsapp_agente_exportar,
    encuesta_whatsapp_agente_detalle_exportar,
    # WhatsApp sin agente
    encuesta_whatsApp_,
    encuesta_whatsapp_detalle,
    encuesta_whatsapp_exportar,
    encuesta_whatsapp_detalle_exportar,
    # Página Web
    encuesta_experiencia_web,
    encuesta_experiencia_web_detalle,
    encuesta_experiencia_web_exportar,
    # Feria Salud
    encuesta_feria_salud_,
    encuesta_feria_salud_exportar,
    # Perfil Agentes
    perfil_agentes_sat_index,
    perfil_agentes_sat_data,
    perfil_agente_sat_detalle,
    perfil_agente_sat_ajax,
    perfil_agentes_od_index,
    perfil_agentes_od_data,
    perfil_agente_od_detalle,
    perfil_agente_od_ajax,
)

# ── Formularios — Tarjetas ────────────────────────────────────────
from .formularios import (
    formulario_tarjetas,
    soli_tarj_credito_buscar,
    soli_tarj_credito_lista,
    soli_tarj_credito_detalle,
    soli_tarj_credito_export,
    soli_tarj_credito_export_detalle,
    soli_tarj_debito_buscar,
    soli_tarj_debito_lista,
    soli_tarj_debito_detalle,
    soli_tarj_debito_export,
    soli_tarj_debito_gestion_buscar,
    soli_tarj_debito_gestion_lista,
    soli_tarj_debito_gestion_detalle,
    soli_tarj_debito_gestion_export,
    soli_redencion_puntos_buscar,
    soli_redencion_puntos_lista,
    soli_redencion_puntos_detalle,
    soli_redencion_puntos_export,
    soli_redencion_puntos_export_detalle,
    caja_ande_asistencia_buscar,
    caja_ande_asistencia_lista,
    caja_ande_asistencia_detalle,
    caja_ande_asistencia_export,
    caja_ande_asistencia_export_detalle,
    # Ahorros
    formulario_ahorros,
    soli_deposito_salario_buscar_cedula,
    soli_deposito_salario_lista,
    soli_deposito_salario_detalle,
    soli_deposito_salario_export,
    soli_deposito_salario_export_detalle,
    soli_ahorro_mod_cuota_buscar,
    soli_ahorro_mod_cuota_lista,
    soli_ahorro_mod_cuota_detalle,
    soli_ahorro_mod_cuota_export,
    soli_ahorro_mod_cuota_export_detalle,
    soli_reinversion_ahorro_buscar,
    soli_reinversion_ahorro_lista,
    soli_reinversion_ahorro_detalle,
    soli_reinversion_ahorro_export,
    soli_reinversion_ahorro_export_detalle,
    soli_autorizacion_ahorro_nuevo_buscar,
    soli_autorizacion_ahorro_nuevo_lista,
    soli_autorizacion_ahorro_nuevo_detalle,
    soli_autorizacion_ahorro_nuevo_export,
    soli_autorizacion_ahorro_nuevo_export_detalle,
    # Vivienda
    formulario_vivienda,
    soli_compra_vehiculo_buscar,
    soli_compra_vehiculo_lista,
    soli_compra_vehiculo_detalle,
    soli_compra_vehiculo_export,
    soli_compra_vehiculo_export_detalle,
    soli_prestamo_vivienda_buscar,
    soli_prestamo_vivienda_lista,
    soli_prestamo_vivienda_detalle,
    soli_prestamo_vivienda_export,
    soli_prestamo_vivienda_export_detalle,
    soli_prestamo_desarrollo_buscar,
    soli_prestamo_desarrollo_lista,
    soli_prestamo_desarrollo_detalle,
    soli_prestamo_desarrollo_export,
    soli_prestamo_desarrollo_export_detalle,
    # Préstamos
    formulario_prestamos,
    soli_presolicitud_credito_personal_buscar,
    soli_presolicitud_credito_personal_lista,
    soli_presolicitud_credito_personal_detalle,
    soli_presolicitud_credito_personal_export,
    soli_presolicitud_credito_personal_export_detalle,
    # Control Crédito
    formulario_control_credito,
    comprobante_autorizacion_ahorro_buscar,
    comprobante_autorizacion_ahorro_lista,
    comprobante_autorizacion_ahorro_detalle,
    comprobante_autorizacion_ahorro_export,
    comprobante_autorizacion_ahorro_export_detalle,
    comprobantes_pago_buscar,
    comprobantes_pago_lista,
    comprobantes_pago_detalle,
    comprobantes_pago_export,
    comprobantes_pago_export_detalle,
    # Servicio Accionista
    formulario_servicio_accionista,
    soli_clave_temporal_cajatel_buscar,
    soli_clave_temporal_cajatel_lista,
    soli_clave_temporal_cajatel_detalle,
    soli_clave_temporal_cajatel_export,
    soli_clave_temporal_cajatel_export_detalle,
    # Seguros
    formulario_seguros,
    soli_seguro_viajero_buscar,
    soli_seguro_viajero_lista,
    soli_seguro_viajero_detalle,
    soli_seguro_viajero_export,
    soli_seguro_viajero_export_detalle,
    soli_marchamo_buscar,
    soli_marchamo_lista,
    soli_marchamo_detalle,
    soli_marchamo_export,
    soli_marchamo_export_detalle,
)

# ── Usuarios ─────────────────────────────────────────────────────
from .usuarios import (
    agentes_home,
    agente_crear,
    agente_editar,
    agente_eliminar,
    agente_qr,
    agente_qr_download,
    agente_qr_download_zip,
    agentes_inactivos,
    agente_restaurar,
)

# ── Logs ─────────────────────────────────────────────────────────
from .logs import (
    logs_home,
    log_detalle,
    logs_export_excel,
    logs_export_pdf,
)