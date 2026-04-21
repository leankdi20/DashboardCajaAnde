# apps/dashboard/reports/soli_compra_vehiculo.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_compra_vehiculo"


class ReporteSolicitudCompraVehiculo:

    QUERY = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto,
            Correo, FechaNacimiento, EstadoCivil,
            TelefonoCelular, TelefonoCasa, TelefonoTrabajo,
            Nacionalidad,
            DireccionDomicilio, ProvinciaDomicilio, CantonDomicilio, DistritoDomicilio,
            DireccionTrabajo, ProvinciaTrabajo, CantonTrabajo, DistritoTrabajo,
            MontoCreditoSolicitado, Plazo, GastosFormalizacion,
            FechaEntrega, TipoVehiculo, Garantia
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto,
            Correo, FechaNacimiento, EstadoCivil,
            TelefonoCelular, TelefonoCasa, TelefonoTrabajo,
            Nacionalidad,
            DireccionDomicilio, ProvinciaDomicilio, CantonDomicilio, DistritoDomicilio,
            DireccionTrabajo, ProvinciaTrabajo, CantonTrabajo, DistritoTrabajo,
            MontoCreditoSolicitado, Plazo, GastosFormalizacion,
            FechaEntrega, TipoVehiculo, Garantia
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto,
            Correo, FechaNacimiento, EstadoCivil,
            TelefonoCelular, TelefonoCasa, TelefonoTrabajo,
            Nacionalidad,
            DireccionDomicilio, ProvinciaDomicilio, CantonDomicilio, DistritoDomicilio,
            DireccionTrabajo, ProvinciaTrabajo, CantonTrabajo, DistritoTrabajo,
            MontoCreditoSolicitado, Plazo, GastosFormalizacion,
            FechaEntrega, TipoVehiculo, Garantia
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                              AS total_solicitudes,
            SUM(CASE WHEN TipoVehiculo = 'Nuevo'  THEN 1 ELSE 0 END)            AS total_nuevo,
            SUM(CASE WHEN TipoVehiculo = 'Usado'  THEN 1 ELSE 0 END)            AS total_usado,
            SUM(CASE WHEN GastosFormalizacion = 'Sí' THEN 1 ELSE 0 END)         AS total_con_gastos,
            SUM(CASE WHEN GastosFormalizacion = 'No' THEN 1 ELSE 0 END)         AS total_sin_gastos,
            AVG(CASE WHEN MontoCreditoSolicitado IS NOT NULL
                     THEN TRY_CAST(MontoCreditoSolicitado AS FLOAT) END)         AS promedio_monto
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_TIPO_VEHICULO = f"""
        SELECT TipoVehiculo, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoVehiculo IS NOT NULL
        GROUP BY TipoVehiculo
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
            condiciones.append("AND RTRIM(NombreCompleto) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("tipo_vehiculo"):
            condiciones.append("AND TipoVehiculo = %s")
            params.append(filtros["tipo_vehiculo"])

        if filtros.get("gastos_formalizacion"):
            condiciones.append("AND GastosFormalizacion = %s")
            params.append(filtros["gastos_formalizacion"])

        if filtros.get("provincia_domicilio"):
            condiciones.append("AND ProvinciaDomicilio = %s")
            params.append(filtros["provincia_domicilio"])

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
    def obtener_tipos_vehiculo(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO_VEHICULO)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])