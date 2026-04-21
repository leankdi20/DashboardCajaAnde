import requests
from typing import Any


class AuthApiServiceLDAP:
    def __init__(self, base_url: str, timeout: int = 15) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

    def validate_user(self, username: str, password: str) -> dict[str, Any] | None:
        """
        Retorna el objeto Login si las credenciales son válidas y tieneAcceso=True.
        Retorna None si las credenciales son inválidas.
        Lanza excepción si hay error de conexión/servidor.
        """
        url = f"{self.base_url}/api/Login/"

        # C# usa JsonConvert.SerializeObject con las propiedades del modelo Login
        # que tiene userName y password (camelCase exacto)
        payload = {
            "userName": username,
            "password": password,
        }

        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        response = requests.post(
            url,
            json=payload,
            headers=headers,
            timeout=self.timeout,
            verify=True,
        )

        response.raise_for_status()

        data = response.json()

        # Replica la lógica de C#: (login != null && login.tieneAcceso) ? login : null
        if data and data.get("tieneAcceso"):
            return data

        return None