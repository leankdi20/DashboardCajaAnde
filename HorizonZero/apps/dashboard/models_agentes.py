# ═══════════════════════════════════════════════════════════════════
# ARCHIVO 1: apps/dashboard/models_agentes.py
# Agrega al final de models.py (o crea este archivo y lo importás)
# ═══════════════════════════════════════════════════════════════════

"""
NO se crean tablas nuevas — la tabla agentes ya existe en SQL Server.
Solo mapeamos el modelo como managed=False para leer/escribir.
Las FK (sucursal_id, unidad_id) se manejan como IntegerField
porque apuntan a tablas de otra BD (encuestas_db).
"""

from django.db import models


class Agente(models.Model):
    """
    Mapea dbo.agentes en la BD 'reportes_db' (encuestas_db).
    managed=False → Django nunca toca la tabla.
    """
    agente_id    = models.AutoField(primary_key=True)
    nombre       = models.CharField(max_length=255)
    sucursal_id  = models.IntegerField()
    unidad_id    = models.IntegerField()
    nombre_lower = models.CharField(max_length=255, blank=True)

    class Meta:
        managed  = False
        db_table = '"encuestas_db"."dbo"."agentes"'   # schema completo para mssql
        # Si tu router usa alias 'reportes_db', cambiá db_table a solo 'agentes'
        # y agregá: app_label = 'dashboard'

    def __str__(self):
        return self.nombre


class Sucursal(models.Model):
    sucursal_id = models.AutoField(primary_key=True)
    nombre      = models.CharField(max_length=255)
    orden       = models.IntegerField(default=0)
    activo      = models.BooleanField(default=True)

    class Meta:
        managed  = False
        db_table = '"encuestas_db"."dbo"."sucursales"'

    def __str__(self):
        return self.nombre


class Unidad(models.Model):
    unidad_id   = models.AutoField(primary_key=True)
    nombre      = models.CharField(max_length=255)
    activo      = models.BooleanField(default=True)
    es_whatsapp = models.BooleanField(default=False)

    class Meta:
        managed  = False
        db_table = '"encuestas_db"."dbo"."unidades"'

    def __str__(self):
        return self.nombre


# ═══════════════════════════════════════════════════════════════════
# PERMISOS — agregar a DashboardAccess.Meta.permissions en models.py
# ═══════════════════════════════════════════════════════════════════
NUEVOS_PERMISOS = [
    ("view_agentes",   "Agentes: Ver listado"),
    ("add_agente",     "Agentes: Crear agente"),
    ("change_agente",  "Agentes: Editar agente"),
    ("delete_agente",  "Agentes: Eliminar agente"),
]
# Copiá estas 4 tuplas dentro de la lista permissions de DashboardAccess