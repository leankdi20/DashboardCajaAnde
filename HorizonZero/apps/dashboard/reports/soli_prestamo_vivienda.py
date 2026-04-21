# apps/dashboard/reports/soli_prestamo_vivienda.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_prestamos_vivienda"


class ReporteSolicitudPrestamoVivienda:

    QUERY = f"""
        SELECT
            respuesta_id, FechaHora, Cedula,
            NombreCompleto, Telefono, TipoPrestamo
        FROM {VISTA}
        ORDER BY FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            respuesta_id, FechaHora, Cedula,
            NombreCompleto, Telefono, TipoPrestamo
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula,
            NombreCompleto, Telefono, TipoPrestamo
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(*)                                                                                AS total_solicitudes,
            SUM(CASE WHEN TipoPrestamo = 'Compra de casa'                       THEN 1 ELSE 0 END) AS total_compra_casa,
            SUM(CASE WHEN TipoPrestamo = 'Mejoras de vivienda'                  THEN 1 ELSE 0 END) AS total_mejoras,
            SUM(CASE WHEN TipoPrestamo = 'Compra de lote y construcción'        THEN 1 ELSE 0 END) AS total_lote_construccion,
            SUM(CASE WHEN TipoPrestamo = 'Construcción de vivienda'             THEN 1 ELSE 0 END) AS total_construccion,
            SUM(CASE WHEN TipoPrestamo = 'Compra de lote'                       THEN 1 ELSE 0 END) AS total_compra_lote,
            SUM(CASE WHEN TipoPrestamo LIKE N'%%Cancelaci%%'                    THEN 1 ELSE 0 END) AS total_cancelaciones,
            SUM(CASE WHEN TipoPrestamo LIKE N'%%Hipotecario%%'                  THEN 1 ELSE 0 END) AS total_hipotecarios,
            SUM(CASE WHEN TipoPrestamo NOT IN (
                'Compra de casa','Mejoras de vivienda','Compra de lote y construcción',
                'Construcción de vivienda','Compra de lote'
            ) AND TipoPrestamo NOT LIKE N'%%Cancelaci%%'
              AND TipoPrestamo NOT LIKE N'%%Hipotecario%%'
              AND TipoPrestamo IS NOT NULL                                      THEN 1 ELSE 0 END) AS total_otros
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
    """

    QUERY_POR_TIPO = f"""
        SELECT TipoPrestamo, COUNT(*) AS total
        FROM {VISTA}
        WHERE TipoPrestamo IS NOT NULL
        GROUP BY TipoPrestamo
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

        if filtros.get("telefono"):
            condiciones.append("AND RTRIM(Telefono) = %s")
            params.append(filtros["telefono"].strip())

        if filtros.get("tipo_prestamo"):
            condiciones.append("AND TipoPrestamo = %s")
            params.append(filtros["tipo_prestamo"])

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
    def obtener_tipos_prestamo(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_POR_TIPO)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])