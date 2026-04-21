# apps/dashboard/reports/soli_marchamo.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_Pago_Marchamo"


class ReporteSolicitudMarchamo:

    QUERY = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto, telefono,
            TipoVehiculo, NumeroPlaca, NombreDuenoRegistral,
            AutorizoPagoAhorro, AutorizoPagoTarjeta, SucursalRetiro,
            SeguroRC, SeguroVida, SeguroSalud, SeguroAsistencia,
            SeguroVidaPlus, SeguroComprensivo
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto, telefono,
            TipoVehiculo, NumeroPlaca, NombreDuenoRegistral,
            AutorizoPagoAhorro, AutorizoPagoTarjeta, SucursalRetiro,
            SeguroRC, SeguroVida, SeguroSalud, SeguroAsistencia,
            SeguroVidaPlus, SeguroComprensivo
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, NombreCompleto, telefono,
            TipoVehiculo, NumeroPlaca, NombreDuenoRegistral,
            AutorizoPagoAhorro, AutorizoPagoTarjeta, SucursalRetiro,
            SeguroRC, SeguroVida, SeguroSalud, SeguroAsistencia,
            SeguroVidaPlus, SeguroComprensivo
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                    AS total_solicitudes,
            SUM(CASE WHEN TipoVehiculo = 'Particular'    THEN 1 ELSE 0 END)            AS total_particular,
            SUM(CASE WHEN TipoVehiculo = 'Moto'          THEN 1 ELSE 0 END)            AS total_moto,
            SUM(CASE WHEN TipoVehiculo = 'Carga liviana' THEN 1 ELSE 0 END)            AS total_carga,
            SUM(CASE WHEN AutorizoPagoAhorro IS NOT NULL
                    AND AutorizoPagoAhorro != ''        THEN 1 ELSE 0 END)            AS total_pago_ahorro,
            SUM(CASE WHEN AutorizoPagoTarjeta IS NOT NULL
                    AND AutorizoPagoTarjeta != ''       THEN 1 ELSE 0 END)            AS total_pago_tarjeta,
            SUM(CASE WHEN SeguroRC   = N'S' + NCHAR(237) THEN 1 ELSE 0 END)            AS total_rc,
            SUM(CASE WHEN SeguroVida = N'S' + NCHAR(237) THEN 1 ELSE 0 END)            AS total_vida
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_TIPO = f"""
        SELECT TipoVehiculo, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoVehiculo IS NOT NULL
        GROUP BY TipoVehiculo
        ORDER BY total DESC
    """

    QUERY_POR_SUCURSAL = f"""
        SELECT SucursalRetiro, COUNT(*) AS total
        FROM {VISTA}
        WHERE SucursalRetiro IS NOT NULL
        GROUP BY SucursalRetiro
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
        if filtros.get("placa"):
            condiciones.append("AND RTRIM(NumeroPlaca) = %s")
            params.append(filtros["placa"].strip())
        if filtros.get("tipo_vehiculo"):
            condiciones.append("AND TipoVehiculo = %s")
            params.append(filtros["tipo_vehiculo"])
        if filtros.get("sucursal"):
            condiciones.append("AND SucursalRetiro = %s")
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
    def obtener_tipos_vehiculo(cls):
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO)

    @classmethod
    def obtener_sucursales(cls):
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_SUCURSAL)

    @classmethod
    def buscar_opciones(cls, columna, termino):
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])