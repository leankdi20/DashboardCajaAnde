import logging
import requests
import time

from django.conf import settings
from django.db import connections


logger = logging.getLogger(__name__)

_api_token = None

_session = requests.Session()
_session.trust_env = False
_session.headers.update({
    "Accept": "application/json",
})


def _get_api_token() -> str:
    global _api_token

    if _api_token:
        return _api_token

    inicio = time.time()

    response = _session.post(
        f"{settings.API_URL}/api/auth/token-interno/",
        json={},
        headers={
            "X-Internal-Key": settings.API_INTERNAL_KEY,
            "Accept": "application/json",
        },
        timeout=30,
    )

    duracion = round(time.time() - inicio, 2)
    logger.warning(f"[DB_SERVICE] token-interno tardó {duracion}s")

    response.raise_for_status()
    _api_token = response.json()["access"]
    return _api_token

def _ejecutar_via_api(sql: str, params: list) -> list[dict]:
    global _api_token

    token = _get_api_token()


    inicio = time.time()

    try:
        response = _session.post(
            f"{settings.API_URL}/api/query/",
            json={"sql": sql, "params": params or []},
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
            },
            timeout=60,
        )

        duracion = round(time.time() - inicio, 2)
        

    except Exception as e:
        duracion = round(time.time() - inicio, 2)
        print(f">>> [DB_SERVICE] /api/query falló después de {duracion}s | ERROR: {e}")
        raise

    if response.status_code == 401:
        _api_token = None
        token = _get_api_token()

        inicio = time.time()

        response = _session.post(
            f"{settings.API_URL}/api/query/",
            json={"sql": sql, "params": params or []},
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
            },
            timeout=60,
        )


    response.raise_for_status()

    inicio_json = time.time()
    data = response.json()
    
    return data


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
        try:
            if getattr(settings, "USE_API", False):
                _ejecutar_via_api(sql, params or [])
                return

            if sql.strip().upper().startswith("WITH"):
                sql = ";" + sql

            with connections[cls.DB_ALIAS].cursor() as cursor:
                cursor.execute(sql, params or [])

        except Exception as e:
            logger.error(f"[ReportesDBService] Comando error: {e}")
            raise