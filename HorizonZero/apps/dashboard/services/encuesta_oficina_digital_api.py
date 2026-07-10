from __future__ import annotations

import json
import time
from collections import defaultdict
from datetime import datetime
from threading import Lock

from django.conf import settings

from apps.dashboard.services.api_client import APIClient


OFFICE_API_URL = getattr(
    settings,
    "ENCUESTAS_OFICINA_DIGITAL_API_URL",
    getattr(settings, "API_URL", ""),
)

LIST_LIMIT = None
LOOKUP_LIMIT = 500
LIST_CACHE_TTL_SECONDS = 300
DETAIL_CACHE_TTL_SECONDS = 900
EXPORT_CACHE_TTL_SECONDS = 300
CACHE_SCHEMA_VERSION = "office_digital_v1"

_list_cache: dict[str, dict] = {}
_list_cache_lock = Lock()
_detail_cache: dict[int, dict] = {}
_detail_cache_lock = Lock()
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

    if all(key in data for key in ("respuesta_id", "encuesta_id")):
        return [data]

    return []


def _extract_detail_payload(data: dict | list | None) -> dict:
    if isinstance(data, dict):
        for key in ("data", "result", "payload"):
            if isinstance(data.get(key), dict):
                return data[key]
        return data
    if isinstance(data, list) and data and isinstance(data[0], dict):
        return data[0]
    return {}


def _score_from_answer(answer: str | None) -> int | None:
    return {
        "Excelente": 3,
        "Regular": 2,
        "Malo": 1,
    }.get((answer or "").strip())


def _classify_from_average(avg: float | None) -> str:
    if avg is None:
        return ""
    if avg >= 2.5:
        return "promotor"
    if avg >= 2:
        return "pasivo"
    return "detractor"


def _cache_key_from_filters(filtros: dict | None = None, include_preguntas: bool = False) -> str:
    payload = {
        "version": CACHE_SCHEMA_VERSION,
        "filtros": filtros or {},
        "include_preguntas": include_preguntas,
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


class EncuestaOficinaDigitalAPIService:
    @classmethod
    def _get(cls, endpoint: str, params: dict | None = None, request=None) -> dict | list:
        return APIClient.get(endpoint, params=params, request=request, base_url=OFFICE_API_URL)

    @classmethod
    def _list_params(cls, filtros: dict | None = None, include_preguntas: bool = False) -> dict:
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
        if include_preguntas:
            params["include_preguntas"] = "true"
        return {k: v for k, v in params.items() if v not in (None, "")}

    @classmethod
    def _fetch_list(cls, filtros: dict | None = None, request=None, include_preguntas: bool = False) -> list[dict]:
        cache_key = _cache_key_from_filters(filtros, include_preguntas=include_preguntas)
        cache = _export_cache if include_preguntas else _list_cache
        lock = _export_cache_lock if include_preguntas else _list_cache_lock
        ttl = EXPORT_CACHE_TTL_SECONDS if include_preguntas else LIST_CACHE_TTL_SECONDS
        cached = _get_ttl_cached(cache, lock, cache_key, ttl)
        if cached is not None:
            return list(cached)

        data = cls._get(
            "encuestas/oficina-digital",
            params=cls._list_params(filtros, include_preguntas=include_preguntas),
            request=request,
        )
        results = _extract_results(data)
        _set_ttl_cached(cache, lock, cache_key, list(results))
        return results

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
            "clasificacion": item.get("clasificacion") or "",
        }

    @classmethod
    def _normalize_detail(cls, item: dict) -> tuple[dict, list[dict], float, str]:
        preguntas = []
        scores = []
        for row in sorted(item.get("preguntas") or [], key=lambda p: p.get("orden") or 0):
            answer = row.get("respuesta") or ""
            preguntas.append({
                "orden": row.get("orden"),
                "encuesta_det_id": row.get("encuesta_det_id"),
                "Pregunta": row.get("pregunta") or "",
                "Respuesta": answer,
            })
            score = _score_from_answer(answer)
            if score is not None:
                scores.append(score)

        promedio = round(sum(scores) / len(scores), 2) if scores else 0
        clasificacion = _classify_from_average(promedio) if scores else ""

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
    def _get_cached_detail(cls, respuesta_id: int):
        return _get_ttl_cached(_detail_cache, _detail_cache_lock, str(respuesta_id), DETAIL_CACHE_TTL_SECONDS)

    @classmethod
    def _set_cached_detail(cls, respuesta_id: int, value: dict):
        _set_ttl_cached(_detail_cache, _detail_cache_lock, str(respuesta_id), value)

    @classmethod
    def obtener_listado(cls, filtros: dict | None = None, request=None) -> list[dict]:
        return [cls._normalize_list_item(item) for item in cls._fetch_list(filtros=filtros, request=request)]

    @classmethod
    def obtener_detalle(cls, respuesta_id: int, request=None) -> dict:
        cached = cls._get_cached_detail(respuesta_id)
        if cached is not None:
            return cached

        item = _extract_detail_payload(cls._get(f"encuestas/oficina-digital/{respuesta_id}", request=request))
        encabezado, preguntas, promedio, clasificacion = cls._normalize_detail(item)
        result = {
            "encabezado": encabezado,
            "preguntas": preguntas,
            "promedio_encuesta": promedio,
            "clasificacion": clasificacion,
        }
        cls._set_cached_detail(respuesta_id, result)
        return result

    @classmethod
    def obtener_detalles_para_exportar(cls, filtros: dict | None = None, request=None) -> list[dict]:
        detalles = []
        for item in cls._fetch_list(filtros=filtros, request=request, include_preguntas=True):
            encabezado, preguntas, promedio, clasificacion = cls._normalize_detail(item)
            if encabezado.get("respuesta_id"):
                cls._set_cached_detail(
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
                    "promedio_encuesta": promedio,
                    "clasificacion": clasificacion,
                })
                continue

            for pregunta in preguntas:
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
                    "Pregunta": pregunta.get("Pregunta", ""),
                    "Respuesta": pregunta.get("Respuesta", ""),
                    "promedio_encuesta": promedio,
                    "clasificacion": clasificacion,
                })
        return detalles

    @classmethod
    def buscar_valores(cls, campo: str, term: str, request=None) -> list[dict]:
        term = (term or "").strip().lower()
        if not term:
            return []

        key = {"agente": "Agente", "nombre": "Nombre"}.get(campo)
        if not key:
            return []

        values = []
        seen = set()
        for item in cls.obtener_listado(request=request)[:LOOKUP_LIMIT]:
            value = (item.get(key) or "").strip()
            normalized = value.lower()
            if not value or term not in normalized or normalized in seen:
                continue
            seen.add(normalized)
            values.append({"id": value, "text": value})
            if len(values) >= 30:
                break
        values.sort(key=lambda row: row["text"])
        return values

    @classmethod
    def obtener_kpis_globales(cls, request=None, fecha_inicio: str | None = None, fecha_fin: str | None = None) -> dict:
        filtros = {
            "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin,
        }
        detalles = cls._fetch_list(filtros=filtros, request=request, include_preguntas=True)
        total = len(detalles)
        promedios = []
        promotores = pasivos = detractores = 0

        for item in detalles:
            _, _, promedio, _ = cls._normalize_detail(item)
            if not promedio:
                continue
            promedios.append(promedio)
            if promedio >= 2.5:
                promotores += 1
            elif promedio >= 2:
                pasivos += 1
            else:
                detractores += 1

        clasificados = promotores + pasivos + detractores
        promedio_general = round((sum(promedios) / len(promedios) / 3) * 100) if promedios else 0

        return {
            "total_encuestas": total,
            "promedio_general": promedio_general,
            "promotores": promotores,
            "promotores_pct": round((promotores / clasificados) * 100) if clasificados else 0,
            "pasivos": pasivos,
            "pasivos_pct": round((pasivos / clasificados) * 100) if clasificados else 0,
            "detractores": detractores,
            "detractores_pct": round((detractores / clasificados) * 100) if clasificados else 0,
        }

    @classmethod
    def obtener_timeline(cls, request=None, filtros: dict | None = None) -> list[dict]:
        timeline = defaultdict(int)
        for item in cls.obtener_listado(filtros=filtros, request=request):
            dt = _parse_iso(item.get("Fecha"))
            if not dt and item.get("Fecha"):
                try:
                    dt = datetime.strptime(item["Fecha"], "%Y-%m-%d")
                except ValueError:
                    dt = None
            if not dt:
                continue
            timeline[(dt.year, dt.month)] += 1

        return [
            {"anio": year, "mes": month, "total": total}
            for (year, month), total in sorted(timeline.items())
        ]

    @classmethod
    def obtener_promedio_agente(cls, agente: str, request=None) -> dict:
        if not agente:
            return {"total_encuestas": 0, "promedio_agente": 0}

        detalles = cls._fetch_list(
            filtros={"agente": agente},
            request=request,
            include_preguntas=True,
        )
        promedios = []
        for item in detalles:
            _, _, promedio, _ = cls._normalize_detail(item)
            if promedio:
                promedios.append(promedio)

        return {
            "total_encuestas": len(detalles),
            "promedio_agente": round(sum(promedios) / len(promedios), 2) if promedios else 0,
        }
