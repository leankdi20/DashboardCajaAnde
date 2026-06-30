# apps/dashboard/services/api_client.py
import requests
from django.conf import settings


class APIClient:
    """
    Cliente HTTP para consumir la API interna.
    Lee el token JWT de la cookie hz_token del request.
    """

    _session = requests.Session()
    _session.trust_env = False
    _session.headers.update({
        "Connection": "close",
        "Accept": "application/json",
    })

    @staticmethod
    def _headers(request=None) -> dict:
        token = ""
        if request:
            token = request.COOKIES.get("hz_token", "")
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Connection": "close",
            "Accept": "application/json",
        }

    @classmethod
    def get(cls, endpoint: str, params: dict = None, request=None) -> dict | list:
        url = f"{settings.API_URL}/api/{endpoint}"
        params = {k: v for k, v in (params or {}).items() if v}
        response = cls._session.get(
            url,
            params=params,
            headers=cls._headers(request),
            timeout=60,
        )
        response.raise_for_status()
        return response.json()

    @classmethod
    def post(cls, endpoint: str, data: dict = None, request=None) -> dict:
        url = f"{settings.API_URL}/api/{endpoint}"
        response = cls._session.post(
            url,
            json=data or {},
            headers=cls._headers(request),
            timeout=60,
        )
        response.raise_for_status()
        return response.json()

    @classmethod
    def put(cls, endpoint: str, data: dict = None, request=None) -> dict:
        url = f"{settings.API_URL}/api/{endpoint}"
        response = cls._session.put(
            url,
            json=data or {},
            headers=cls._headers(request),
            timeout=60,
        )
        response.raise_for_status()
        return response.json()

    @classmethod
    def delete(cls, endpoint: str, request=None) -> dict:
        url = f"{settings.API_URL}/api/{endpoint}"
        response = cls._session.delete(
            url,
            headers=cls._headers(request),
            timeout=60,
        )
        response.raise_for_status()
        return response.json()