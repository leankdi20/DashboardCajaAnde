# apps/dashboard/reports/comprobante_autorizacion_ahorro.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Comprobante_autorizacion_ahorro"


class ReporteComprobanteAutorizacionAhorro:

    QUERY = f"""
        SELECT
            respuesta_id, Fecha, Hora, Cedula, NombreCompleto, NumeroTelefonico,
            DetallePago,
            Documento1, Documento2, Documento3, Documento4, Documento5,
            Documento6, Documento7, Documento8, Documento9, Documento10
        FROM {VISTA}
        ORDER BY Fecha DESC, Hora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id, Fecha, Hora, Cedula, NombreCompleto, NumeroTelefonico,
            DetallePago,
            Documento1, Documento2, Documento3, Documento4, Documento5,
            Documento6, Documento7, Documento8, Documento9, Documento10
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY Fecha DESC, Hora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, Fecha, Hora, Cedula, NombreCompleto, NumeroTelefonico,
            DetallePago,
            Documento1, Documento2, Documento3, Documento4, Documento5,
            Documento6, Documento7, Documento8, Documento9, Documento10
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*) AS total_comprobantes,
            SUM(CASE WHEN Documento1 IS NOT NULL THEN 1 ELSE 0 END) AS total_con_doc1,
            SUM(CASE WHEN Documento2 IS NOT NULL THEN 1 ELSE 0 END) AS total_con_doc2,
            MAX(
                CASE WHEN Documento1  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento2  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento3  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento4  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento5  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento6  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento7  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento8  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento9  IS NOT NULL THEN 1 ELSE 0 END +
                CASE WHEN Documento10 IS NOT NULL THEN 1 ELSE 0 END
            ) AS max_documentos
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
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
        condiciones, params = [], []
        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())
        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(NombreCompleto) = %s")
            params.append(filtros["nombre"].strip())
        if filtros.get("fecha_inicio"):
            condiciones.append("AND Fecha >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones.append("AND Fecha <= %s")
            params.append(filtros["fecha_fin"])
        return " ".join(condiciones), params

    @classmethod
    def obtener_datos(cls, filtros=None):
        if not filtros or not any(filtros.values()):
            return ReportesDBService.ejecutar_query(cls.QUERY)
        sql_filtros, params = cls._construir_filtros(filtros)
        return ReportesDBService.ejecutar_query(
            cls.QUERY_FILTRADO.format(filtros_base=sql_filtros), params)

    @classmethod
    def obtener_detalle(cls, respuesta_id):
        r = ReportesDBService.ejecutar_query(cls.QUERY_DETALLE, [respuesta_id])
        return r[0] if r else None

    @classmethod
    def obtener_kpis(cls, filtros=None):
        sql_filtros, params = cls._construir_filtros(filtros or {})
        r = ReportesDBService.ejecutar_query(
            cls.QUERY_KPIS.format(filtros_base=sql_filtros), params)
        return r[0] if r else {}

    @classmethod
    def buscar_opciones(cls, columna, termino):
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])