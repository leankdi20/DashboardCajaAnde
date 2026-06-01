import logging
import requests
from django.conf import settings
from django.db import connections

logger = logging.getLogger(__name__)

_api_token = None


def _get_api_token() -> str:
    global _api_token
    if _api_token:
        return _api_token
    response = requests.post(
        f"{settings.API_URL}/api/token/",
        json={
            "username": settings.API_USER,
            "password": settings.API_PASSWORD,
        },
        timeout=10,
    )
    response.raise_for_status()
    _api_token = response.json()["access"]
    return _api_token


def _ejecutar_via_api(sql: str, params: list) -> list[dict]:
    global _api_token
    token = _get_api_token()

    response = requests.post(
        f"{settings.API_URL}/api/query/",
        json={"sql": sql, "params": params or []},
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )

    if response.status_code == 401:
        _api_token = None
        token = _get_api_token()
        response = requests.post(
            f"{settings.API_URL}/api/query/",
            json={"sql": sql, "params": params or []},
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )

    response.raise_for_status()
    return response.json()


class ReportesDBService:
    DB_ALIAS = "default"

    @classmethod
    def ejecutar_query(cls, sql, params=None) -> list[dict]:
        try:
            if getattr(settings, "USE_API", False):
                return _ejecutar_via_api(sql, params or [])

            if sql.strip().upper().startswith("WITH"):
                sql = ";" + sql

            with connections[cls.DB_ALIAS].cursor() as cursor:
                cursor.execute(sql, params or [])
                columnas = [col[0] for col in cursor.description]
                return [dict(zip(columnas, fila)) for fila in cursor.fetchall()]

        except Exception as e:
            logger.error(f"[ReportesDBService] Error: {e}")
            raise

    @classmethod
    def ejecutar_comando(cls, sql: str, params=None) -> None:
        """Para INSERT/UPDATE/DELETE que no retornan filas."""
        try:
            if getattr(settings, "USE_API", False):
                # Enviar via API — reutiliza el mismo endpoint /api/query/
                _ejecutar_via_api(sql, params or [])
                return

            if sql.strip().upper().startswith("WITH"):
                sql = ";" + sql

            with connections[cls.DB_ALIAS].cursor() as cursor:
                cursor.execute(sql, params or [])

        except Exception as e:
            logger.error(f"[ReportesDBService] Comando error: {e}")
            raise