# apps/dashboard/services/api_client.py
import requests
from requests.adapters import HTTPAdapter
from django.conf import settings

from apps.usuarios.session_control import APISessionExpiredError


class APIClient:
    """
    Cliente HTTP para consumir la API interna.
    Lee el token JWT de la cookie hz_token del request.
    """

    _session = requests.Session()
    _session.trust_env = False
    _session.mount("http://", HTTPAdapter(pool_connections=50, pool_maxsize=50, max_retries=0))
    _session.mount("https://", HTTPAdapter(pool_connections=50, pool_maxsize=50, max_retries=0))
    _session.headers.update({
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
            "Accept": "application/json",
        }

    @staticmethod
    def _raise_if_session_invalid(response):
        if response.status_code not in (401, 403):
            return

        message = "Su sesion fue iniciada en otro navegador o dispositivo."
        try:
            payload = response.json()
            message = (
                payload.get("detail")
                or payload.get("message")
                or message
            )
        except ValueError:
            pass

        raise APISessionExpiredError(message=message)

    @classmethod
    def get(cls, endpoint: str, params: dict = None, request=None, base_url: str = None) -> dict | list:
        root = (base_url or settings.API_URL).rstrip("/")
        url = f"{root}/api/{endpoint.lstrip('/')}"
        params = {k: v for k, v in (params or {}).items() if v}
        response = cls._session.get(
            url,
            params=params,
            headers=cls._headers(request),
            timeout=60,
        )
        cls._raise_if_session_invalid(response)
        response.raise_for_status()
        return response.json()

    @classmethod
    def post(cls, endpoint: str, data: dict = None, request=None, base_url: str = None) -> dict:
        root = (base_url or settings.API_URL).rstrip("/")
        url = f"{root}/api/{endpoint.lstrip('/')}"
        response = cls._session.post(
            url,
            json=data or {},
            headers=cls._headers(request),
            timeout=60,
        )
        cls._raise_if_session_invalid(response)
        response.raise_for_status()
        return response.json()

    @classmethod
    def put(cls, endpoint: str, data: dict = None, request=None, base_url: str = None) -> dict:
        root = (base_url or settings.API_URL).rstrip("/")
        url = f"{root}/api/{endpoint.lstrip('/')}"
        response = cls._session.put(
            url,
            json=data or {},
            headers=cls._headers(request),
            timeout=60,
        )
        cls._raise_if_session_invalid(response)
        response.raise_for_status()
        return response.json()

    @classmethod
    def delete(cls, endpoint: str, request=None, base_url: str = None) -> dict:
        root = (base_url or settings.API_URL).rstrip("/")
        url = f"{root}/api/{endpoint.lstrip('/')}"
        response = cls._session.delete(
            url,
            headers=cls._headers(request),
            timeout=60,
        )
        cls._raise_if_session_invalid(response)
        response.raise_for_status()
        return response.json()
