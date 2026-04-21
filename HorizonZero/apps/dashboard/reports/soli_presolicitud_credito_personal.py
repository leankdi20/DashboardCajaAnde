# apps/dashboard/reports/soli_presolicitud_credito_personal.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_presolicitud_credito_personal"


class ReporteSolicitudPresolicitudCreditoPersonal:

    QUERY = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto,
            Telefono, TipoCredito, SucursalFormalizacion, Monto,
            FrenteCedula_URL, ReversoCedula_URL, DesglosePension_URL
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto,
            Telefono, TipoCredito, SucursalFormalizacion, Monto,
            FrenteCedula_URL, ReversoCedula_URL, DesglosePension_URL
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto,
            Telefono, TipoCredito, SucursalFormalizacion, Monto,
            FrenteCedula_URL, ReversoCedula_URL, DesglosePension_URL
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                AS total_solicitudes,
            SUM(CASE WHEN TipoCredito = 'Corriente'                  THEN 1 ELSE 0 END) AS total_corriente,
            SUM(CASE WHEN TipoCredito = 'Alternativo personal'       THEN 1 ELSE 0 END) AS total_alternativo,
            SUM(CASE WHEN TipoCredito = 'Especial'                   THEN 1 ELSE 0 END) AS total_especial,
            SUM(CASE WHEN TipoCredito = 'Alternativo para cancelación de deudas'
                                                                     THEN 1 ELSE 0 END) AS total_cancelacion,
            SUM(CASE WHEN TipoCredito = 'Salud'                      THEN 1 ELSE 0 END) AS total_salud,
            SUM(CASE WHEN TipoCredito NOT IN (
                'Corriente','Alternativo personal','Especial',
                'Alternativo para cancelación de deudas','Salud'
            ) AND TipoCredito IS NOT NULL                            THEN 1 ELSE 0 END) AS total_otros,
            AVG(CASE WHEN Monto IS NOT NULL
                     THEN TRY_CAST(Monto AS FLOAT) END)                                 AS promedio_monto,
            SUM(CASE WHEN FrenteCedula_URL   IS NOT NULL THEN 1 ELSE 0 END)             AS con_frente,
            SUM(CASE WHEN DesglosePension_URL IS NOT NULL THEN 1 ELSE 0 END)            AS con_desglose
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_TIPO = f"""
        SELECT TipoCredito, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoCredito IS NOT NULL
        GROUP BY TipoCredito
        ORDER BY total DESC
    """

    QUERY_POR_SUCURSAL = f"""
        SELECT SucursalFormalizacion, COUNT(*) AS total
        FROM {VISTA}
        WHERE SucursalFormalizacion IS NOT NULL
        GROUP BY SucursalFormalizacion
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
        condiciones, params = [], []
        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())
        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(NombreCompleto) = %s")
            params.append(filtros["nombre"].strip())
        if filtros.get("tipo_credito"):
            condiciones.append("AND TipoCredito = %s")
            params.append(filtros["tipo_credito"])
        if filtros.get("sucursal"):
            condiciones.append("AND SucursalFormalizacion = %s")
            params.append(filtros["sucursal"])
        if filtros.get("fecha_inicio"):
            condiciones.append("AND CAST(FechaHora AS DATE) >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones.append("AND CAST(FechaHora AS DATE) <= %s")
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
    def obtener_tipos_credito(cls):
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO)

    @classmethod
    def obtener_sucursales(cls):
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_SUCURSAL)

    @classmethod
    def buscar_opciones(cls, columna, termino):
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])