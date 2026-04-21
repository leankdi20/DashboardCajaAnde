import logging
from django.db import connections

logger = logging.getLogger(__name__)


class ReportesDBService:
    DB_ALIAS = "reportes_db"

    @classmethod
    def ejecutar_query(cls, sql, params=None) -> list[dict]:
        try:
            with connections[cls.DB_ALIAS].cursor() as cursor:
                # SQL Server requiere ; antes de WITH cuando hay sesión activa
                if sql.strip().upper().startswith("WITH"):
                    sql = ";" + sql
                cursor.execute(sql, params or [])
                columnas = [col[0] for col in cursor.description]
                return [dict(zip(columnas, fila)) for fila in cursor.fetchall()]
        except Exception as e:
            print(f"Error ejecutando query: {e}")
            raise
        
    @classmethod
    def ejecutar_comando(cls, sql: str) -> None:
        """Para INSERT/UPDATE/DELETE que no retornan filas."""
        try:
            with connections[cls.DB_ALIAS].cursor() as cursor:
                cursor.execute(sql)
        except Exception as e:
            print(f"Error ejecutando comando: {e}")
            raise

# class Reportes2DBService:
#     DB_ALIAS = "reportes_db2"  # Usuario/contraseña para producción