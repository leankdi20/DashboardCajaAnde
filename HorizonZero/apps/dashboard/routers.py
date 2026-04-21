# ═══════════════════════════════════════════════════════════════════
# ARCHIVO 2: apps/dashboard/routers.py  
# ═══════════════════════════════════════════════════════════════════
"""
Router que dirige los modelos Agente/Sucursal/Unidad a 'reportes_db'.
Agregá en settings.py:
    DATABASE_ROUTERS = ['apps.dashboard.routers.EncuestasDBRouter']
"""

MODELOS_ENCUESTAS = {'agente', 'sucursal', 'unidad'}


class EncuestasDBRouter:
    def db_for_read(self, model, **hints):
        if model.__name__.lower() in MODELOS_ENCUESTAS:
            return 'reportes_db'
        return None

    def db_for_write(self, model, **hints):
        if model.__name__.lower() in MODELOS_ENCUESTAS:
            return 'reportes_db'
        return None

    def allow_relation(self, obj1, obj2, **hints):
        return None

    def allow_migrate(self, db, app_label, model_name=None, **hints):
        # Nunca migrar estos modelos
        if model_name and model_name.lower() in MODELOS_ENCUESTAS:
            return False
        return None