# apps/dashboard/reports/soli_autorizacion_ahorro_nuevo.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_Autorizacion_Ahorro_Nuevo"


class ReporteSolicitudAutorizacionAhorroNuevo:

    QUERY = f"""
        SELECT
            respuesta_id,
            Fecha,
            Cedula,
            Nombre_Completo,
            TipoAhorro,
            FormaPago,
            TipoReinversion,
            RetiroIntereses,
            MontoCuota,
            MontoDebitarAhorro
        FROM {VISTA}
        ORDER BY Fecha DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id,
            Fecha,
            Cedula,
            Nombre_Completo,
            TipoAhorro,
            FormaPago,
            TipoReinversion,
            RetiroIntereses,
            MontoCuota,
            MontoDebitarAhorro
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY Fecha DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id,
            Fecha,
            Cedula,
            Nombre_Completo,
            TipoAhorro,
            FormaPago,
            TipoReinversion,
            RetiroIntereses,
            MontoCuota,
            MontoDebitarAhorro
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                            AS total_solicitudes,
            -- Por forma de pago
            SUM(CASE WHEN FormaPago = 'Cuota'         THEN 1 ELSE 0 END)                       AS total_cuota,
            SUM(CASE WHEN FormaPago = 'Monto inicial' THEN 1 ELSE 0 END)                       AS total_monto_inicial,
            SUM(CASE WHEN FormaPago = 'Mixto'         THEN 1 ELSE 0 END)                       AS total_mixto,
            -- Por tipo de reinversión
            SUM(CASE WHEN TipoReinversion LIKE N'%%Capital m%%' THEN 1 ELSE 0 END)             AS total_cap_intereses,
            SUM(CASE WHEN TipoReinversion = 'Reinversión solo capital' THEN 1 ELSE 0 END)      AS total_solo_capital,
            SUM(CASE WHEN TipoReinversion = 'Reinversión cuota'        THEN 1 ELSE 0 END)      AS total_reinv_cuota,
            SUM(CASE WHEN TipoReinversion = 'Sin reinversión'          THEN 1 ELSE 0 END)      AS total_sin_reinversion,
            -- Montos
            AVG(CASE WHEN MontoCuota IS NOT NULL
                     THEN CAST(MontoCuota AS FLOAT) END)                                       AS promedio_cuota,
            SUM(CASE WHEN MontoCuota IS NOT NULL THEN 1 ELSE 0 END)                            AS total_con_cuota,
            SUM(CASE WHEN MontoDebitarAhorro IS NOT NULL THEN 1 ELSE 0 END)                    AS total_con_debito
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

    QUERY_POR_FORMA_PAGO = f"""
        SELECT FormaPago, COUNT(*) AS total
        FROM {VISTA}
        WHERE FormaPago IS NOT NULL
        GROUP BY FormaPago
        ORDER BY total DESC
    """

    QUERY_POR_TIPO_REINVERSION = f"""
        SELECT TipoReinversion, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoReinversion IS NOT NULL
        GROUP BY TipoReinversion
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

        if filtros.get("forma_pago"):
            condiciones.append("AND FormaPago = %s")
            params.append(filtros["forma_pago"])

        if filtros.get("tipo_reinversion"):
            condiciones.append("AND TipoReinversion = %s")
            params.append(filtros["tipo_reinversion"])

        if filtros.get("retiro_intereses"):
            condiciones.append("AND RetiroIntereses = %s")
            params.append(filtros["retiro_intereses"])

        if filtros.get("fecha_inicio"):
            condiciones.append("AND Fecha >= %s")
            params.append(filtros["fecha_inicio"])

        if filtros.get("fecha_fin"):
            condiciones.append("AND Fecha <= %s")
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
    def obtener_formas_pago(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_FORMA_PAGO)

    @classmethod
    def obtener_tipos_reinversion(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO_REINVERSION)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])