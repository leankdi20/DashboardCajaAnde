# apps/dashboard/reports/soli_tarj_debito_gestion.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_solicitud_tarjeta_debito"


class ReporteSolicitudTarjetaDebitoGestion:

    # ── Una fila por accionista (pivotamos orden 1 y 2) ──────────
    QUERY = f"""
        SELECT
            v1.respuesta_id,
            v1.FechaHora,
            v1.Cedula,
            v1.Nombre,
            v1.Telefono,
            v1.Correo,
            v1.Respuesta  AS TipoTramite,
            v2.Respuesta  AS DireccionEnvio
        FROM {VISTA} v1
        LEFT JOIN {VISTA} v2
            ON v1.respuesta_id = v2.respuesta_id
            AND v2.orden = 2
        WHERE v1.orden = 1
        ORDER BY v1.FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            v1.respuesta_id,
            v1.FechaHora,
            v1.Cedula,
            v1.Nombre,
            v1.Telefono,
            v1.Correo,
            v1.Respuesta  AS TipoTramite,
            v2.Respuesta  AS DireccionEnvio
        FROM {VISTA} v1
        LEFT JOIN {VISTA} v2
            ON v1.respuesta_id = v2.respuesta_id
            AND v2.orden = 2
        WHERE v1.orden = 1
        {{filtros_base}}
        ORDER BY v1.FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, Nombre,
            Telefono, Correo, orden, Pregunta, Respuesta
        FROM {VISTA}
        WHERE respuesta_id = %s
        ORDER BY orden
    """

    # ── KPIs por tipo de trámite ─────────────────────────────────
    QUERY_KPIS = f"""
        SELECT
            COUNT(DISTINCT respuesta_id)                                                               AS total_solicitudes,
            SUM(CASE WHEN Respuesta = 'Renovación de tarjeta de débito'   THEN 1 ELSE 0 END)          AS total_renovacion,
            SUM(CASE WHEN Respuesta = 'Solicitud de tarjeta de débito nueva' THEN 1 ELSE 0 END)       AS total_nueva,
            SUM(CASE WHEN Respuesta = 'Estado de cuenta de tarjeta de débito' THEN 1 ELSE 0 END)      AS total_estado_cuenta,
            SUM(CASE WHEN Respuesta = 'Reposición por deterioro'          THEN 1 ELSE 0 END)          AS total_reposicion,
            SUM(CASE WHEN Respuesta = 'Saldo del ahorro'                  THEN 1 ELSE 0 END)          AS total_saldo_ahorro,
            SUM(CASE WHEN Respuesta = 'Reclamos'                          THEN 1 ELSE 0 END)          AS total_reclamos
        FROM {VISTA}
        WHERE orden = 1
        {{filtros_base}}
    """

    # ── Opciones para filtros (Select2 estático) ─────────────────
    QUERY_TIPOS_TRAMITE = f"""
        SELECT Respuesta AS TipoTramite, COUNT(*) AS total
        FROM {VISTA}
        WHERE orden = 1 AND Respuesta IS NOT NULL
        GROUP BY Respuesta
        ORDER BY total DESC
    """

    QUERY_DESTINOS = f"""
        SELECT Respuesta AS Destino, COUNT(*) AS total
        FROM {VISTA}
        WHERE orden = 2 AND Respuesta IS NOT NULL
        GROUP BY Respuesta
        ORDER BY total DESC
    """

    # ── Búsqueda AJAX ────────────────────────────────────────────
    QUERY_BUSCAR_OPCIONES = """
        SELECT DISTINCT TOP 30 RTRIM({col}) AS {col}
        FROM {vista}
        WHERE {col} IS NOT NULL
        AND RTRIM({col}) LIKE %s
        ORDER BY RTRIM({col})
    """

    # ─────────────────────────────────────────────
    @staticmethod
    def _construir_filtros(filtros: dict):
        condiciones = []
        params = []

        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(v1.Cedula) = %s")
            params.append(filtros["cedula"].strip())

        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(v1.Nombre) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("telefono"):
            condiciones.append("AND RTRIM(v1.Telefono) = %s")
            params.append(filtros["telefono"].strip())

        if filtros.get("correo"):
            condiciones.append("AND RTRIM(v1.Correo) = %s")
            params.append(filtros["correo"].strip())

        if filtros.get("tipo_tramite"):
            condiciones.append("AND v1.Respuesta = %s")
            params.append(filtros["tipo_tramite"])

        if filtros.get("destino"):
            condiciones.append("AND v2.Respuesta = %s")
            params.append(filtros["destino"])

        if filtros.get("fecha_inicio"):
            condiciones.append("AND CAST(v1.FechaHora AS DATE) >= %s")
            params.append(filtros["fecha_inicio"])

        if filtros.get("fecha_fin"):
            condiciones.append("AND CAST(v1.FechaHora AS DATE) <= %s")
            params.append(filtros["fecha_fin"])

        return " ".join(condiciones), params

    # ── _construir_filtros simple (sin alias v1) para KPIS ───────
    @staticmethod
    def _construir_filtros_simple(filtros: dict):
        """Para queries que no usan JOIN (KPIS, tipos, destinos)."""
        condiciones = []
        params = []

        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())

        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(Nombre) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("tipo_tramite"):
            condiciones.append("AND Respuesta = %s")
            params.append(filtros["tipo_tramite"])

        if filtros.get("fecha_inicio"):
            condiciones.append("AND CAST(FechaHora AS DATE) >= %s")
            params.append(filtros["fecha_inicio"])

        if filtros.get("fecha_fin"):
            condiciones.append("AND CAST(FechaHora AS DATE) <= %s")
            params.append(filtros["fecha_fin"])

        return " ".join(condiciones), params

    # ─────────────────────────────────────────────
    @classmethod
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        if not filtros or not any(filtros.values()):
            return ReportesDBService.ejecutar_query(cls.QUERY)
        sql_filtros, params = cls._construir_filtros(filtros)
        sql = cls.QUERY_FILTRADO.format(filtros_base=sql_filtros)
        return ReportesDBService.ejecutar_query(sql, params)

    @classmethod
    def obtener_detalle(cls, respuesta_id: int) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_DETALLE, [respuesta_id])

    @classmethod
    def obtener_kpis(cls, filtros: dict = None) -> dict:
        sql_filtros, params = cls._construir_filtros_simple(filtros or {})
        sql = cls.QUERY_KPIS.format(filtros_base=sql_filtros)
        resultado = ReportesDBService.ejecutar_query(sql, params)
        return resultado[0] if resultado else {}

    @classmethod
    def obtener_tipos_tramite(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_TIPOS_TRAMITE)

    @classmethod
    def obtener_destinos(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_DESTINOS)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])