# apps/dashboard/reports/soli_tarj_credito.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_solicitud_tarjeta_credito"


class ReporteSolicitudTarjetaCredito:

    QUERY = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, Nombre,
            Telefono, Correo, TipoTramite, DireccionEnvio,
            URL_Cedula_Frente, URL_Cedula_Reverso
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, Nombre,
            Telefono, Correo, TipoTramite, DireccionEnvio,
            URL_Cedula_Frente, URL_Cedula_Reverso
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, Nombre,
            Telefono, Correo, TipoTramite, DireccionEnvio,
            URL_Cedula_Frente, URL_Cedula_Reverso
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                             AS total_solicitudes,
            SUM(CASE WHEN TipoTramite LIKE N'%%tarjeta nueva%%'             THEN 1 ELSE 0 END)  AS total_nuevas,
            SUM(CASE WHEN TipoTramite LIKE N'%%Renovaci%%'                  THEN 1 ELSE 0 END)  AS total_renovacion,
            SUM(CASE WHEN TipoTramite LIKE N'%%Estado de cuenta%%'          THEN 1 ELSE 0 END)  AS total_estado_cuenta,
            SUM(CASE WHEN TipoTramite LIKE N'%%ampliaci%%'                  THEN 1 ELSE 0 END)  AS total_ampliacion,
            SUM(CASE WHEN TipoTramite LIKE N'%%Compra de saldos de%%'       THEN 1 ELSE 0 END)  AS total_compra_saldos,
            SUM(CASE WHEN TipoTramite LIKE N'%%Reposici%%deterioro%%'       THEN 1 ELSE 0 END)  AS total_reposicion,
            SUM(CASE WHEN TipoTramite LIKE N'%%Reclamo%%'                   THEN 1 ELSE 0 END)  AS total_reclamos,
            SUM(CASE WHEN TipoTramite LIKE N'%%sin tarjeta activa%%'        THEN 1 ELSE 0 END)  AS total_saldo_sin_tarjeta
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_TRAMITE = f"""
        SELECT TipoTramite, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoTramite IS NOT NULL
        GROUP BY TipoTramite
        ORDER BY total DESC
    """

    # ── AJAX: búsqueda de opciones por campo ─────
    # Trae máximo 30 resultados que contengan el término buscado
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
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())

        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(Nombre) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("telefono"):
            condiciones.append("AND RTRIM(Telefono) = %s")
            params.append(filtros["telefono"].strip())

        if filtros.get("correo"):
            condiciones.append("AND RTRIM(Correo) = %s")
            params.append(filtros["correo"].strip())

        if filtros.get("tipo_tramite"):
            condiciones.append("AND TipoTramite = %s")
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
    def obtener_detalle(cls, respuesta_id: int) -> dict | None:
        resultado = ReportesDBService.ejecutar_query(cls.QUERY_DETALLE, [respuesta_id])
        return resultado[0] if resultado else None

    @classmethod
    def obtener_kpis(cls, filtros: dict = None) -> dict:
        sql_filtros, params = cls._construir_filtros(filtros or {})
        sql = cls.QUERY_KPIS.format(filtros_base=sql_filtros)
        resultado = ReportesDBService.ejecutar_query(sql, params)
        return resultado[0] if resultado else {}

    @classmethod
    def obtener_por_tramite(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TRAMITE)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        """Para Select2 AJAX — devuelve hasta 30 coincidencias."""
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])