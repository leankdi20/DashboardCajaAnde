# apps/dashboard/reports/soli_tarj_debito.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Solicitud_tarjeta_débito_Ciudadano_Oro"


class ReporteSolicitudTarjetaDebito:

    # ─────────────────────────────────────────────
    #  Esta vista es tipo encuesta: una fila por pregunta/respuesta.
    #  Agrupamos por respuesta_id para mostrar una fila por accionista.
    # ─────────────────────────────────────────────

    QUERY = f"""
        SELECT respuesta_id, FechaHora, Cedula, Nombre, Correo,
            [Teléfono] AS Telefono, Respuesta AS DireccionEnvio
        FROM {VISTA}
        WHERE Pregunta = '¿Dónde desea que enviemos su tarjeta?'
        ORDER BY FechaHora DESC
    """


    QUERY_FILTRADO = f"""
        SELECT respuesta_id, FechaHora, Cedula, Nombre, Correo,
            [Teléfono] AS Telefono, Respuesta AS DireccionEnvio
        FROM {VISTA}
        WHERE Pregunta = '¿Dónde desea que enviemos su tarjeta?'
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
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(DISTINCT respuesta_id)                                                                AS total_solicitudes,
            SUM(CASE WHEN Respuesta = 'San José (oficinas centrales)'         THEN 1 ELSE 0 END)        AS total_san_jose,
            SUM(CASE WHEN Respuesta = 'Correo (entrega en casa de habitación)' THEN 1 ELSE 0 END)       AS total_correo,
            SUM(CASE WHEN Respuesta = 'Heredia'                               THEN 1 ELSE 0 END)        AS total_heredia,
            SUM(CASE WHEN Respuesta = 'Alajuela'                              THEN 1 ELSE 0 END)        AS total_alajuela,
            SUM(CASE WHEN Respuesta = 'Cartago'                               THEN 1 ELSE 0 END)        AS total_cartago,
            SUM(CASE WHEN Respuesta = 'Liberia'                               THEN 1 ELSE 0 END)        AS total_liberia,
            SUM(CASE WHEN Respuesta = 'Santa Cruz'                            THEN 1 ELSE 0 END)        AS total_santa_cruz,
            SUM(CASE WHEN Respuesta = 'San Ramón'                             THEN 1 ELSE 0 END)        AS total_san_ramon,
            SUM(CASE WHEN Respuesta = 'Puntarenas'                            THEN 1 ELSE 0 END)        AS total_puntarenas,
            SUM(CASE WHEN Respuesta = 'San Carlos'                            THEN 1 ELSE 0 END)        AS total_san_carlos,
            SUM(CASE WHEN Respuesta = 'Pérez Zeledón'                         THEN 1 ELSE 0 END)        AS total_perez_zeledon,
            SUM(CASE WHEN Respuesta = 'Guápiles'                              THEN 1 ELSE 0 END)        AS total_guapiles,
            SUM(CASE WHEN Respuesta = 'Ciudad Neily'                          THEN 1 ELSE 0 END)        AS total_ciudad_neily,
            SUM(CASE WHEN Respuesta = 'Limón'                                 THEN 1 ELSE 0 END)        AS total_limon
        FROM {VISTA}
        WHERE Pregunta = '¿Dónde desea que enviemos su tarjeta?'
        {{filtros_base}}
    """


    QUERY_POR_DESTINO = f"""
        SELECT Respuesta AS Destino, COUNT(*) AS total
        FROM {VISTA}
        WHERE Pregunta = '¿Dónde desea que enviemos su tarjeta?'
        AND Respuesta IS NOT NULL
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

        if filtros.get("correo"):
            condiciones.append("AND RTRIM(Correo) = %s")
            params.append(filtros["correo"].strip())

        if filtros.get("telefono"):
            condiciones.append("AND RTRIM([Teléfono]) = %s")
            params.append(filtros["telefono"].strip())

        if filtros.get("correo"):
            condiciones.append("AND RTRIM(Correo) = %s")
            params.append(filtros["correo"].strip())

        if filtros.get("destino"):
            condiciones.append("AND Respuesta = %s")
            params.append(filtros["destino"])

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
    def obtener_por_destino(cls) -> list[dict]:
        sql = cls.QUERY_POR_DESTINO.format(filtros_base="")
        return ReportesDBService.ejecutar_query(sql)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])