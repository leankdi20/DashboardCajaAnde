from __future__ import annotations

import json
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from threading import Lock

import requests
from django.conf import settings

from apps.dashboard.services.api_client import APIClient


DEV_API_URL = getattr(
    settings,
    "ENCUESTAS_SATISFACCION_API_URL",
    "http://172.16.21.204:9080",
)

LIST_LIMIT = None
LOOKUP_LIMIT = 500
DEBUG_DUMP_PATH = Path(settings.BASE_DIR) / "encuesta_satisfaccion_debug.json"
MAX_429_RETRIES = 6
DETAIL_CACHE_TTL_SECONDS = 900
LIST_CACHE_TTL_SECONDS = 300
EXPORT_CACHE_TTL_SECONDS = 300
CACHE_SCHEMA_VERSION = "sat_export_v4"

_detail_cache: dict[int, dict] = {}
_detail_cache_lock = Lock()
_list_cache: dict[str, dict] = {}
_list_cache_lock = Lock()
_export_cache: dict[str, dict] = {}
_export_cache_lock = Lock()


def _parse_iso(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _format_fecha(value: str | None) -> str:
    dt = _parse_iso(value)
    return dt.strftime("%Y-%m-%d") if dt else (value or "")


def _format_hora(value: str | None) -> str:
    dt = _parse_iso(value)
    return dt.strftime("%H:%M:%S") if dt else (value or "")


def _date_only(value: str | None) -> str | None:
    dt = _parse_iso(value)
    if dt:
        return dt.strftime("%Y-%m-%d")
    if value and len(value) >= 10:
        return value[:10]
    return None


def _extract_results(data: dict | list | None) -> list[dict]:
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if not isinstance(data, dict):
        return []

    for key in ("results", "resultados", "data", "items"):
        value = data.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        if isinstance(value, dict):
            nested = _extract_results(value)
            if nested:
                return nested

    for key in ("result", "payload", "response"):
        value = data.get(key)
        if isinstance(value, dict):
            nested = _extract_results(value)
            if nested:
                return nested
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]

    if all(key in data for key in ("respuesta_id", "encuesta_id")):
        return [data]

    return []


def _extract_detail_payload(data: dict | list | None) -> dict:
    if isinstance(data, dict):
        if isinstance(data.get("data"), dict):
            return data["data"]
        if isinstance(data.get("result"), dict):
            return data["result"]
        if isinstance(data.get("payload"), dict):
            return data["payload"]
        return data

    if isinstance(data, list) and data and isinstance(data[0], dict):
        return data[0]

    return {}


def _score_from_answer(question: str, answer: str) -> int | None:
    question = (question or "").lower()
    answer = (answer or "").strip()

    if "satisfecho" in question or "califica su experiencia" in question:
        return {
            "Muy satisfecho": 5,
            "Satisfecho": 4,
            "Ni satisfecho ni insatisfecho": 3,
            "Poco satisfecho": 2,
            "Nada satisfecho": 1,
        }.get(answer)

    if "como fue su experiencia" in question:
        return {
            "Muy facil": 5,
            "Muy fácil": 5,
            "Facil": 4,
            "Fácil": 4,
            "Ni facil ni dificil": 3,
            "Ni fácil ni difícil": 3,
            "Dificil": 2,
            "Difícil": 2,
            "Muy dificil": 1,
            "Muy difícil": 1,
        }.get(answer)

    return None


def _classify_from_average(avg: float | None) -> str:
    if avg is None:
        return ""
    if avg >= 4:
        return "promotor"
    if avg == 3:
        return "pasivo"
    if avg <= 2:
        return "detractor"
    return ""


def _normalize_recommendation(question: str, answer: str) -> str:
    if "recomienda" not in (question or "").lower():
        return answer or ""
    value = str(answer or "").strip()
    if value == "1":
        return "Sí"
    if value in {"0", "2"}:
        return "No"
    return answer or ""


def _normalize_timeline_item(item: dict) -> dict:
    return {
        "anio": item.get("anio", item.get("year", item.get("anio_num"))),
        "mes": item.get("mes", item.get("month", item.get("mes_num"))),
        "total": item.get("total", item.get("count", item.get("cantidad", 0))),
    }


def _safe_debug_value(value):
    if isinstance(value, dict):
        return {str(k): _safe_debug_value(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_safe_debug_value(v) for v in value[:5]]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _debug_dump(section: str, endpoint: str, params: dict | None, raw_data):
    try:
        payload = {}
        if DEBUG_DUMP_PATH.exists():
            payload = json.loads(DEBUG_DUMP_PATH.read_text(encoding="utf-8"))

        payload[section] = {
            "timestamp": datetime.now().isoformat(),
            "endpoint": endpoint,
            "params": _safe_debug_value(params or {}),
            "raw_type": type(raw_data).__name__,
            "raw_preview": _safe_debug_value(raw_data),
        }
        DEBUG_DUMP_PATH.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _debug_dump_error(section: str, endpoint: str, params: dict | None, error: Exception):
    try:
        payload = {}
        if DEBUG_DUMP_PATH.exists():
            payload = json.loads(DEBUG_DUMP_PATH.read_text(encoding="utf-8"))

        response = getattr(error, "response", None)
        payload[f"{section}__error"] = {
            "timestamp": datetime.now().isoformat(),
            "endpoint": endpoint,
            "params": _safe_debug_value(params or {}),
            "error_type": type(error).__name__,
            "error_message": str(error),
            "status_code": getattr(response, "status_code", None),
            "response_text": getattr(response, "text", None),
        }
        DEBUG_DUMP_PATH.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _retry_delay_from_response(response, attempt: int) -> float:
    if response is None:
        return min(2 ** attempt, 30)

    retry_after = response.headers.get("Retry-After")
    if retry_after:
        try:
            return max(float(retry_after), 1.0)
        except ValueError:
            pass

    text = getattr(response, "text", "") or ""
    match = re.search(r"available in (\d+) seconds", text, re.IGNORECASE)
    if match:
        return max(float(match.group(1)), 1.0)

    return min(2 ** attempt, 30)


def _get_cached_detail(respuesta_id: int):
    now = time.time()
    with _detail_cache_lock:
        cached = _detail_cache.get(respuesta_id)
        if not cached:
            return None
        if now - cached["stored_at"] > DETAIL_CACHE_TTL_SECONDS:
            _detail_cache.pop(respuesta_id, None)
            return None
        return cached["value"]


def _set_cached_detail(respuesta_id: int, value: dict):
    with _detail_cache_lock:
        _detail_cache[respuesta_id] = {
            "stored_at": time.time(),
            "value": value,
        }


def _cache_key_from_filters(filtros: dict | None = None, limit: int | None = None) -> str:
    payload = {
        "version": CACHE_SCHEMA_VERSION,
        "filtros": filtros or {},
        "limit": limit,
    }
    return json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str)


def _get_ttl_cached(cache: dict, lock: Lock, key: str, ttl_seconds: int):
    now = time.time()
    with lock:
        cached = cache.get(key)
        if not cached:
            return None
        if now - cached["stored_at"] > ttl_seconds:
            cache.pop(key, None)
            return None
        return cached["value"]


def _set_ttl_cached(cache: dict, lock: Lock, key: str, value):
    with lock:
        cache[key] = {
            "stored_at": time.time(),
            "value": value,
        }


class EncuestaSatisfaccionAPIService:
    @classmethod
    def _get(cls, endpoint: str, params: dict | None = None, request=None) -> dict | list:
        section = endpoint.replace("/", "_")
        for attempt in range(MAX_429_RETRIES + 1):
            try:
                data = APIClient.get(endpoint, params=params, request=request, base_url=DEV_API_URL)
                _debug_dump(section, endpoint, params, data)
                return data
            except requests.RequestException as error:
                response = getattr(error, "response", None)
                if getattr(response, "status_code", None) == 429 and attempt < MAX_429_RETRIES:
                    time.sleep(_retry_delay_from_response(response, attempt))
                    continue
                _debug_dump_error(section, endpoint, params, error)
                raise
            except Exception as error:
                _debug_dump_error(section, endpoint, params, error)
                raise

    @classmethod
    def _list_params(
        cls,
        filtros: dict | None = None,
        limit: int | None = LIST_LIMIT,
        include_preguntas: bool = False,
    ) -> dict:
        filtros = filtros or {}
        params = {
            "desde": filtros.get("fecha_inicio"),
            "hasta": filtros.get("fecha_fin"),
            "sucursal": filtros.get("sucursal"),
            "unidad": filtros.get("unidad"),
            "agente": filtros.get("agente"),
            "cedula": filtros.get("cedula"),
            "nombre": filtros.get("nombre"),
            "encuesta_id": filtros.get("encuesta_id"),
            "respuesta_id": filtros.get("respuesta_id"),
            "clasificacion": filtros.get("clasificacion"),
        }
        if limit is not None:
            params["limit"] = limit
        if include_preguntas:
            params["include_preguntas"] = "true"
        return {k: v for k, v in params.items() if v not in (None, "")}

    @classmethod
    def _fetch_all_list_results(
        cls,
        filtros: dict | None = None,
        request=None,
        limit: int | None = LIST_LIMIT,
        include_preguntas: bool = False,
    ) -> list[dict]:
        filtros = filtros or {}
        cache_key = _cache_key_from_filters(
            {**filtros, "__include_preguntas": include_preguntas},
            limit=limit,
        )
        cached = _get_ttl_cached(_list_cache, _list_cache_lock, cache_key, LIST_CACHE_TTL_SECONDS)
        if cached is not None:
            return list(cached)

        base_params = cls._list_params(filtros, limit=limit, include_preguntas=include_preguntas)

        first_page = cls._get("encuestas/satisfaccion", params=base_params, request=request)
        if not isinstance(first_page, dict):
            results = _extract_results(first_page)
            _set_ttl_cached(_list_cache, _list_cache_lock, cache_key, list(results))
            return results

        results = _extract_results(first_page)
        total_count = first_page.get("count")

        if limit is not None and results:
            all_results = list(results)
            seen_ids = {item.get("respuesta_id") for item in all_results if item.get("respuesta_id")}
            offset = len(results)
            expected_total = total_count if isinstance(total_count, int) and total_count > 0 else None

            while True:
                if expected_total is not None and offset >= expected_total:
                    break

                page_params = dict(base_params)
                page_params["offset"] = offset
                page = cls._get("encuestas/satisfaccion", params=page_params, request=request)
                page_results = _extract_results(page)

                if not page_results:
                    break

                new_items = []
                for item in page_results:
                    respuesta_id = item.get("respuesta_id")
                    if respuesta_id and respuesta_id in seen_ids:
                        continue
                    if respuesta_id:
                        seen_ids.add(respuesta_id)
                    new_items.append(item)

                if not new_items:
                    break

                all_results.extend(new_items)
                offset += len(page_results)

                if len(page_results) < limit:
                    break

            if len(all_results) > len(results):
                _set_ttl_cached(_list_cache, _list_cache_lock, cache_key, list(all_results))
                return all_results

        # Fallback: el endpoint real capea count=500 sin metadata de paginacion.
        # Recorremos por ventanas de fecha hacia atras usando el dia mas antiguo recibido.
        all_results = list(results)
        seen_ids = {item.get("respuesta_id") for item in all_results if item.get("respuesta_id")}
        fecha_inicio = filtros.get("fecha_inicio")
        oldest_date = min(
            (_date_only(item.get("fecha")) for item in results if item.get("fecha")),
            default=None,
        )
        previous_oldest = None

        while limit is not None and results and len(results) >= limit and oldest_date and oldest_date != previous_oldest:
            previous_oldest = oldest_date
            next_hasta_dt = _parse_iso(f"{oldest_date}T00:00:00") - timedelta(days=1)
            next_hasta = next_hasta_dt.strftime("%Y-%m-%d")

            if fecha_inicio and next_hasta < fecha_inicio:
                break

            page_params = dict(base_params)
            page_params.pop("offset", None)
            page_params["hasta"] = next_hasta
            page = cls._get("encuestas/satisfaccion", params=page_params, request=request)
            results = _extract_results(page)

            if not results:
                break

            new_items = []
            for item in results:
                respuesta_id = item.get("respuesta_id")
                if respuesta_id and respuesta_id in seen_ids:
                    continue
                if respuesta_id:
                    seen_ids.add(respuesta_id)
                new_items.append(item)

            if not new_items:
                break

            all_results.extend(new_items)
            oldest_date = min(
                (_date_only(item.get("fecha")) for item in results if item.get("fecha")),
                default=None,
            )

        _set_ttl_cached(_list_cache, _list_cache_lock, cache_key, list(all_results))
        return all_results

    @classmethod
    def _normalize_list_item(cls, item: dict) -> dict:
        return {
            "respuesta_id": item.get("respuesta_id"),
            "encuesta_id": item.get("encuesta_id"),
            "Fecha": _format_fecha(item.get("fecha")),
            "Hora": _format_hora(item.get("fecha")),
            "Cedula": item.get("cedula") or "",
            "Nombre": item.get("nombre") or "",
            "Correo": item.get("correo") or "",
            "Agente": item.get("agente") or "",
            "Sucursal": item.get("sucursal") or "",
            "Unidad": item.get("unidad") or "",
            "Gestion": item.get("gestion") or "",
            "encuesta_nombre": item.get("encuesta_nombre") or "",
            "total_preguntas": item.get("total_preguntas") or 0,
        }

    @classmethod
    def _normalize_detail(cls, item: dict) -> tuple[dict, list[dict], float, str]:
        preguntas = []
        scores = []

        for row in sorted(item.get("preguntas") or [], key=lambda p: p.get("orden") or 0):
            question = row.get("pregunta") or ""
            answer = _normalize_recommendation(question, row.get("respuesta") or "")
            preguntas.append({
                "orden": row.get("orden"),
                "encuesta_det_id": row.get("encuesta_det_id"),
                "Pregunta": question,
                "Respuesta": answer,
            })
            score = _score_from_answer(question, row.get("respuesta") or "")
            if score is not None:
                scores.append(score)

        promedio = round(sum(scores) / len(scores), 2) if scores else 0
        clasificacion = _classify_from_average(promedio)

        encabezado = {
            "respuesta_id": item.get("respuesta_id"),
            "encuesta_id": item.get("encuesta_id"),
            "encuesta_nombre": item.get("encuesta_nombre") or "",
            "Fecha": _format_fecha(item.get("fecha")),
            "Hora": _format_hora(item.get("hora") or item.get("fecha")),
            "Cedula": item.get("cedula") or "",
            "Nombre": item.get("nombre") or "",
            "Correo": item.get("correo") or "",
            "Agente": item.get("agente") or "",
            "Sucursal": item.get("sucursal") or "",
            "Unidad": item.get("unidad") or "",
            "Gestion": item.get("gestion") or "",
            "total_preguntas": len(preguntas),
            "promedio_encuesta": promedio,
            "clasificacion": clasificacion,
        }
        return encabezado, preguntas, promedio, clasificacion

    @classmethod
    def obtener_listado(cls, filtros: dict | None = None, request=None) -> list[dict]:
        filtros = (filtros or {}).copy()

        items = [cls._normalize_list_item(item) for item in cls._fetch_all_list_results(filtros=filtros, request=request)]

        gestion = (filtros.get("gestion") or "").strip().lower()
        if gestion:
            items = [item for item in items if gestion in item.get("Gestion", "").lower()]

        return items

    @classmethod
    def obtener_detalle(cls, respuesta_id: int, request=None) -> dict:
        cached = _get_cached_detail(respuesta_id)
        if cached is not None:
            return cached

        item = _extract_detail_payload(cls._get(f"encuestas/satisfaccion/{respuesta_id}", request=request))
        encabezado, preguntas, promedio, clasificacion = cls._normalize_detail(item)
        result = {
            "encabezado": encabezado,
            "preguntas": preguntas,
            "promedio_encuesta": promedio,
            "clasificacion": clasificacion,
        }
        _set_cached_detail(respuesta_id, result)
        return result

    @classmethod
    def obtener_kpis_globales(
        cls,
        request=None,
        sucursales: list | None = None,
        fecha_inicio: str | None = None,
        fecha_fin: str | None = None,
        unidad: str | None = None,
        agente: str | None = None,
    ) -> dict:
        params = {
            "desde": fecha_inicio,
            "hasta": fecha_fin,
            "unidad": unidad,
            "agente": agente,
        }
        if sucursales:
            params["sucursal"] = sucursales[0] if len(sucursales) == 1 else ",".join(sucursales)

        data = cls._get("encuestas/satisfaccion/kpis", params=params, request=request)
        return {
            "total_encuestas": data.get("total_encuestas", 0),
            "promedio_general": (data.get("satisfaccion") or {}).get("valor", 0),
            "indicador_dificultad": (data.get("indice_dificultad") or {}).get("valor", 0),
            "lealtad_pct": (data.get("indice_lealtad") or {}).get("valor", 0),
            "promotores": (data.get("promotores") or {}).get("total_respuestas", 0),
            "promotores_pct": (data.get("promotores") or {}).get("valor", 0),
            "pasivos": (data.get("pasivos") or {}).get("total_respuestas", 0),
            "pasivos_pct": (data.get("pasivos") or {}).get("valor", 0),
            "detractores": (data.get("detractores") or {}).get("total_respuestas", 0),
            "detractores_pct": (data.get("detractores") or {}).get("valor", 0),
            "periodo": data.get("periodo") or {},
        }

    @classmethod
    def obtener_timeline(cls, filtros: dict | None = None, request=None) -> list[dict]:
        filtros = filtros or {}
        params = {
            "desde": filtros.get("fecha_inicio"),
            "hasta": filtros.get("fecha_fin"),
            "sucursal": filtros.get("sucursal"),
            "unidad": filtros.get("unidad"),
            "agente": filtros.get("agente"),
        }
        data = cls._get("encuestas/satisfaccion/timeline", params=params, request=request)
        return [
            row for row in (_normalize_timeline_item(item) for item in _extract_results(data))
            if row.get("anio") and row.get("mes")
        ]

    @classmethod
    def obtener_opciones(
        cls,
        filtros: dict | None = None,
        request=None,
        items: list[dict] | None = None,
    ) -> tuple[list[dict], list[dict]]:
        items = items if items is not None else cls.obtener_listado(filtros=filtros, request=request)
        sucursales = sorted({item.get("Sucursal", "").strip() for item in items if item.get("Sucursal", "").strip()})
        unidades = sorted({item.get("Unidad", "").strip() for item in items if item.get("Unidad", "").strip()})
        return (
            [{"Sucursal": value} for value in sucursales],
            [{"Unidad": value} for value in unidades],
        )

    @classmethod
    def buscar_valores(cls, campo: str, term: str, request=None) -> list[dict]:
        term = (term or "").strip().lower()
        if not term:
            return []

        items = cls._fetch_all_list_results(
            filtros={},
            request=request,
            limit=LOOKUP_LIMIT,
        )

        field_map = {
            "agente": "agente",
            "gestion": "gestion",
            "nombre": "nombre",
        }
        key = field_map.get(campo)
        if not key:
            return []

        values = []
        seen = set()
        for item in items:
            value = (item.get(key) or "").strip()
            if not value:
                continue
            normalized = value.lower()
            if term not in normalized or normalized in seen:
                continue
            seen.add(normalized)
            values.append({"id": value, "text": value})
            if len(values) >= 30:
                break

        values.sort(key=lambda row: row["text"])
        return values

    @classmethod
    def obtener_detalles_para_exportar(cls, filtros: dict | None = None, request=None) -> list[dict]:
        filtros = filtros or {}
        export_cache_key = _cache_key_from_filters(filtros, limit=None)
        cached = _get_ttl_cached(_export_cache, _export_cache_lock, export_cache_key, EXPORT_CACHE_TTL_SECONDS)
        if cached is not None:
            return list(cached)

        data = cls._get(
            "encuestas/satisfaccion",
            params=cls._list_params(filtros, limit=None, include_preguntas=True),
            request=request,
        )
        items = _extract_results(data)
        detalles = []

        for item in items:
            encabezado, preguntas, promedio, clasificacion = cls._normalize_detail(item)
            if encabezado.get("respuesta_id"):
                _set_cached_detail(
                    encabezado["respuesta_id"],
                    {
                        "encabezado": encabezado,
                        "preguntas": preguntas,
                        "promedio_encuesta": promedio,
                        "clasificacion": clasificacion,
                    },
                )

            if not preguntas:
                detalles.append({
                    "respuesta_id": encabezado.get("respuesta_id"),
                    "Fecha": encabezado.get("Fecha", ""),
                    "Hora": encabezado.get("Hora", ""),
                    "Cedula": encabezado.get("Cedula", ""),
                    "Nombre": encabezado.get("Nombre", ""),
                    "Agente": encabezado.get("Agente", ""),
                    "Sucursal": encabezado.get("Sucursal", ""),
                    "Unidad": encabezado.get("Unidad", ""),
                    "Gestion": encabezado.get("Gestion", ""),
                    "Pregunta": "",
                    "Respuesta": "",
                    "orden": None,
                    "promedio_encuesta": promedio,
                    "clasificacion": clasificacion,
                })
                continue

            for pregunta in preguntas:
                detalles.append({
                    "respuesta_id": encabezado["respuesta_id"],
                    "Fecha": encabezado.get("Fecha", ""),
                    "Hora": encabezado.get("Hora", ""),
                    "Cedula": encabezado.get("Cedula", ""),
                    "Nombre": encabezado.get("Nombre", ""),
                    "Agente": encabezado.get("Agente", ""),
                    "Sucursal": encabezado.get("Sucursal", ""),
                    "Unidad": encabezado.get("Unidad", ""),
                    "Gestion": encabezado.get("Gestion", ""),
                    "Pregunta": pregunta["Pregunta"],
                    "Respuesta": pregunta["Respuesta"],
                    "orden": pregunta.get("orden"),
                    "promedio_encuesta": promedio,
                    "clasificacion": clasificacion,
                })
        _set_ttl_cached(_export_cache, _export_cache_lock, export_cache_key, list(detalles))
        return detalles
