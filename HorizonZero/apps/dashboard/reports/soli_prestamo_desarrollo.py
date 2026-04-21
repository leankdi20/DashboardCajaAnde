# apps/dashboard/reports/soli_prestamo_desarrollo.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_prestamo_desarrollo_economico"


class ReporteSolicitudPrestamoDesarrollo:

    QUERY = f"""
        SELECT respuesta_id, FechaHora, Cedula, NombreCompleto, Telefono, PlanInversion
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT respuesta_id, FechaHora, Cedula, NombreCompleto, Telefono, PlanInversion
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT respuesta_id, FechaHora, Cedula, NombreCompleto, Telefono, PlanInversion
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                        AS total_solicitudes,
            SUM(CASE WHEN PlanInversion = 'Comercio'                          THEN 1 ELSE 0 END) AS total_comercio,
            SUM(CASE WHEN PlanInversion = 'Servicios'                         THEN 1 ELSE 0 END) AS total_servicios,
            SUM(CASE WHEN PlanInversion = N'Ganader' + NCHAR(237) + 'a'      THEN 1 ELSE 0 END) AS total_ganaderia,
            SUM(CASE WHEN PlanInversion = 'Turismo'                           THEN 1 ELSE 0 END) AS total_turismo,
            SUM(CASE WHEN PlanInversion = 'Agricultura'                       THEN 1 ELSE 0 END) AS total_agricultura,
            SUM(CASE WHEN PlanInversion = 'Transporte'                        THEN 1 ELSE 0 END) AS total_transporte,
            SUM(CASE WHEN PlanInversion = 'Industria'                         THEN 1 ELSE 0 END) AS total_industria
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_PLAN = f"""
        SELECT PlanInversion, COUNT(*) AS total
        FROM {VISTA}
        WHERE PlanInversion IS NOT NULL
        GROUP BY PlanInversion
        ORDER BY total DESC
    """

    QUERY_BUSCAR_OPCIONES = """
        SELECT DISTINCT TOP 30 RTRIM({col}) AS {col}
        FROM {vista}
        WHERE {col} IS NOT NULL
        AND RTRIM({col}) LIKE %s
        ORDER BY RTRIM({col})
    """

    @staticmethod
    def _construir_filtros(filtros: dict):
        condiciones = []
        params = []
        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())
        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(NombreCompleto) = %s")
            params.append(filtros["nombre"].strip())
        if filtros.get("telefono"):
            condiciones.append("AND RTRIM(Telefono) = %s")
            params.append(filtros["telefono"].strip())
        if filtros.get("plan_inversion"):
            condiciones.append("AND PlanInversion = %s")
            params.append(filtros["plan_inversion"])
        if filtros.get("fecha_inicio"):
            condiciones.append("AND CAST(FechaHora AS DATE) >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones.append("AND CAST(FechaHora AS DATE) <= %s")
            params.append(filtros["fecha_fin"])
        return " ".join(condiciones), params

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
    def obtener_planes(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_PLAN)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])