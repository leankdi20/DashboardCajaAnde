# apps/dashboard/mixins.py
"""
Mixin reutilizable para aplicar restricciones de PerfilUsuario
a cualquier vista de encuestas o formularios.

Uso en cualquier vista:
    filtros, unidad_forzada, sucursal_forzada = aplicar_restricciones_perfil(request, filtros)
"""


def aplicar_restricciones_perfil(request, filtros: dict) -> tuple[dict, str | None, str | None]:
    """
    Lee el PerfilUsuario del request.user y fuerza los filtros
    de unidad y/o sucursal si corresponde.

    Retorna:
        filtros            → dict actualizado con las restricciones
        unidad_forzada     → str o None
        sucursal_forzada   → str o None
    """
    perfil = getattr(request.user, "perfil", None)

    unidad_forzada   = perfil.unidad   if perfil and perfil.unidad   else None
    sucursal_forzada = perfil.sucursal if perfil and perfil.sucursal else None

    if unidad_forzada:
        filtros["unidad"] = unidad_forzada

    if sucursal_forzada:
        filtros["sucursal"] = sucursal_forzada

    return filtros, unidad_forzada, sucursal_forzada