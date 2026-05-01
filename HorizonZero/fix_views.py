import os

base = r'D:\Proyectos\DashboardCajaAnde\HorizonZero\apps\dashboard\views'

base_py = """import json

MESES_ES = {
    1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr',
    5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago',
    9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic',
}


def build_timeline_heatmap(raw_timeline: list) -> tuple:
    timeline_data = [
        {"label": f"{MESES_ES[r['mes']]} {r['anio']}", "total": r["total"]}
        for r in raw_timeline
    ]
    heatmap_dict  = {}
    heatmap_anios = []
    for r in raw_timeline:
        a, m, t = r["anio"], r["mes"], r["total"]
        if a not in heatmap_dict:
            heatmap_dict[a] = {}
            heatmap_anios.append(a)
        heatmap_dict[a][m] = t
    return timeline_data, heatmap_anios, json.dumps(heatmap_dict)
"""

formularios_init = """from .tarjetas import (
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
)
from .ahorros import (
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
)
from .vivienda import (
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
)
from .prestamos import (
    formulario_prestamos,
    soli_presolicitud_credito_personal_buscar,
    soli_presolicitud_credito_personal_lista,
    soli_presolicitud_credito_personal_detalle,
    soli_presolicitud_credito_personal_export,
    soli_presolicitud_credito_personal_export_detalle,
)
from .control_credito import (
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
)
from .servicio_accionista import (
    formulario_servicio_accionista,
    soli_clave_temporal_cajatel_buscar,
    soli_clave_temporal_cajatel_lista,
    soli_clave_temporal_cajatel_detalle,
    soli_clave_temporal_cajatel_export,
    soli_clave_temporal_cajatel_export_detalle,
)
from .seguros import (
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
"""

files = {
    '_base.py': base_py,
    os.path.join('formularios', '__init__.py'): formularios_init,
}

for fname, content in files.items():
    path = os.path.join(base, fname)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'OK: {fname} ({len(content)} bytes)')