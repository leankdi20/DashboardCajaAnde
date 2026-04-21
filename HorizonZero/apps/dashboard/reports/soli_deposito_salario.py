# apps/dashboard/reports/soli_deposito_salario.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_deposito_salario"


class ReporteSolicitudDepositoSalario:

    QUERY = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Cedula,
            BoletaSolicitud_URL,
            FrenteCedula_URL,
            ReversoCedula_URL
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Cedula,
            BoletaSolicitud_URL,
            FrenteCedula_URL,
            ReversoCedula_URL
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id,
            FechaHora,
            Cedula,
            BoletaSolicitud_URL,
            FrenteCedula_URL,
            ReversoCedula_URL
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                   AS total_solicitudes,
            SUM(CASE WHEN BoletaSolicitud_URL IS NOT NULL THEN 1 ELSE 0 END) AS total_con_boleta,
            SUM(CASE WHEN FrenteCedula_URL    IS NOT NULL THEN 1 ELSE 0 END) AS total_con_frente,
            SUM(CASE WHEN ReversoCedula_URL   IS NOT NULL THEN 1 ELSE 0 END) AS total_con_reverso
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_BUSCAR_CEDULA = f"""
        SELECT DISTINCT TOP 30 RTRIM(Cedula) AS Cedula
        FROM {VISTA}
        WHERE Cedula IS NOT NULL
        AND RTRIM(Cedula) LIKE %s
        ORDER BY RTRIM(Cedula)
    """

    # ─────────────────────────────────────────────
    @staticmethod
    def _construir_filtros(filtros: dict):
        condiciones = []
        params = []

        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())

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
    def buscar_cedulas(cls, termino: str) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_BUSCAR_CEDULA, [f"%{termino}%"])