# apps/dashboard/reports/soli_ahorro_mod_cuota.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_Ahorro_modificacion_cuota"


class ReporteSolicitudAhorroModCuota:

    QUERY = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Nombre_Completo,
            Cedula,
            TipoAhorro,
            NumeroContrato,
            TipoModificacion,
            MontoCuotaDeducir
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Nombre_Completo,
            Cedula,
            TipoAhorro,
            NumeroContrato,
            TipoModificacion,
            MontoCuotaDeducir
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Nombre_Completo,
            Cedula,
            TipoAhorro,
            NumeroContrato,
            TipoModificacion,
            MontoCuotaDeducir
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                AS total_solicitudes,
            SUM(CASE WHEN TipoModificacion = 'Aumento'     THEN 1 ELSE 0 END)     AS total_aumentos,
            SUM(CASE WHEN TipoModificacion = 'Disminución' THEN 1 ELSE 0 END)     AS total_disminuciones,
            SUM(CASE WHEN TipoModificacion = 'Aumento'
                    THEN CAST(MontoCuotaDeducir AS FLOAT) ELSE 0 END)
            -
            SUM(CASE WHEN TipoModificacion = 'Disminución'
                    THEN CAST(MontoCuotaDeducir AS FLOAT) ELSE 0 END)             AS impacto_neto,
            AVG(CASE WHEN TipoModificacion = 'Aumento'
                    THEN CAST(MontoCuotaDeducir AS FLOAT) END)                    AS promedio_aumentos,
            AVG(CASE WHEN TipoModificacion = 'Disminución'
                    THEN CAST(MontoCuotaDeducir AS FLOAT) END)                    AS promedio_disminuciones
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_TIPO_AHORRO = f"""
        SELECT TipoAhorro, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoAhorro IS NOT NULL
        GROUP BY TipoAhorro
        ORDER BY total DESC
    """

    QUERY_POR_TIPO_MOD = f"""
        SELECT TipoModificacion, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoModificacion IS NOT NULL
        GROUP BY TipoModificacion
        ORDER BY total DESC
    """

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
            condiciones.append("AND RTRIM(Nombre_Completo) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("tipo_ahorro"):
            condiciones.append("AND TipoAhorro = %s")
            params.append(filtros["tipo_ahorro"])

        if filtros.get("tipo_modificacion"):
            condiciones.append("AND TipoModificacion = %s")
            params.append(filtros["tipo_modificacion"])

        if filtros.get("numero_contrato"):
            condiciones.append("AND RTRIM(NumeroContrato) = %s")
            params.append(filtros["numero_contrato"].strip())

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
    def obtener_tipos_ahorro(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO_AHORRO)

    @classmethod
    def obtener_tipos_modificacion(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO_MOD)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])