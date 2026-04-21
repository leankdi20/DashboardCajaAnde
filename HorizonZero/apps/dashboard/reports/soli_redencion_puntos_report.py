# apps/dashboard/reports/soli_redencion_puntos.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_Redencion_Puntos"


class ReporteSolicitudRedencionPuntos:

    QUERY = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Cedula,
            Nombre,
            Correo,
            [Teléfono]  AS Telefono,
            Respuesta   AS TipoRedencion
        FROM {VISTA}
        WHERE orden = 1
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Cedula,
            Nombre,
            Correo,
            [Teléfono]  AS Telefono,
            Respuesta   AS TipoRedencion
        FROM {VISTA}
        WHERE orden = 1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Cedula,
            Nombre,
            Correo,
            [Teléfono]  AS Telefono,
            Pregunta,
            Respuesta
        FROM {VISTA}
        WHERE respuesta_id = %s
        ORDER BY orden
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(DISTINCT respuesta_id)                                                                         AS total_solicitudes,
            SUM(CASE WHEN Respuesta LIKE N'%%Cash back%%'          THEN 1 ELSE 0 END)                           AS total_cashback,
            SUM(CASE WHEN Respuesta LIKE N'%%Pago de contado%%'    THEN 1 ELSE 0 END)                           AS total_pago_contado,
            SUM(CASE WHEN Respuesta LIKE N'%%Pago m%%nimo%%'       THEN 1 ELSE 0 END)                           AS total_pago_minimo
        FROM {VISTA}
        WHERE orden = 1
        {{filtros_base}}
    """

    QUERY_POR_TIPO = f"""
        SELECT Respuesta AS TipoRedencion, COUNT(*) AS total
        FROM {VISTA}
        WHERE orden = 1 AND Respuesta IS NOT NULL
        {{filtros_base}}
        GROUP BY Respuesta
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
            condiciones.append("AND RTRIM(Nombre) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("telefono"):
            condiciones.append("AND RTRIM([Teléfono]) = %s")
            params.append(filtros["telefono"].strip())

        if filtros.get("correo"):
            condiciones.append("AND RTRIM(Correo) = %s")
            params.append(filtros["correo"].strip())

        if filtros.get("tipo_redencion"):
            condiciones.append("AND Respuesta = %s")
            params.append(filtros["tipo_redencion"])

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
        sql_filtros, params = cls._construir_filtros(filtros or {})
        sql = cls.QUERY_KPIS.format(filtros_base=sql_filtros)
        resultado = ReportesDBService.ejecutar_query(sql, params)
        return resultado[0] if resultado else {}

    @classmethod
    def obtener_por_tipo(cls) -> list[dict]:
        sql = cls.QUERY_POR_TIPO.format(filtros_base="")
        return ReportesDBService.ejecutar_query(sql)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])