# apps/dashboard/reports/soli_caja_ande_asistencia.py

from ..services.db_service import ReportesDBService

VISTA = "dbo.vw_Caja_de_ANDE_Asistencia"


class ReporteCajaAndeAsistencia:

    # ── Una fila por accionista (pivot orden 1 y orden 3) ────────
    QUERY = f"""
        SELECT
            v1.respuesta_id,
            v1.FechaHora,
            v1.Cedula,
            v1.Nombre,
            v1.Correo,
            v1.Respuesta  AS TipoPlan,
            v3.Respuesta  AS TipoTarjeta
        FROM {VISTA} v1
        LEFT JOIN {VISTA} v3
            ON v1.respuesta_id = v3.respuesta_id
            AND v3.orden = 3
        WHERE v1.orden = 1
        ORDER BY v1.FechaHora DESC
    """

    QUERY_FILTRADO = f"""
        SELECT
            v1.respuesta_id,
            v1.FechaHora,
            v1.Cedula,
            v1.Nombre,
            v1.Correo,
            v1.Respuesta  AS TipoPlan,
            v3.Respuesta  AS TipoTarjeta
        FROM {VISTA} v1
        LEFT JOIN {VISTA} v3
            ON v1.respuesta_id = v3.respuesta_id
            AND v3.orden = 3
        WHERE v1.orden = 1
        {{filtros_base}}
        ORDER BY v1.FechaHora DESC
    """

    QUERY_DETALLE = f"""
        SELECT
            respuesta_id, FechaHora, Cedula, Nombre,
            Correo, orden, Pregunta, Respuesta
        FROM {VISTA}
        WHERE respuesta_id = %s
        ORDER BY orden
    """

    QUERY_KPIS = f"""
        SELECT
            COUNT(DISTINCT respuesta_id)                                                                         AS total_solicitudes,
            SUM(CASE WHEN Respuesta LIKE N'%%2.035%%'      THEN 1 ELSE 0 END)                                   AS total_basico,
            SUM(CASE WHEN Respuesta LIKE N'%%Familiar%%'   THEN 1 ELSE 0 END)                                   AS total_familiar,
            SUM(CASE WHEN Respuesta LIKE N'%%Salud%%'      THEN 1 ELSE 0 END)                                   AS total_salud,
            SUM(CASE WHEN Respuesta LIKE N'%%Extranjero%%' THEN 1 ELSE 0 END)                                   AS total_extranjero,
            SUM(CASE WHEN Respuesta LIKE N'%%Edad de Oro%%' THEN 1 ELSE 0 END)                                  AS total_edad_oro
        FROM {VISTA}
        WHERE orden = 1
        {{filtros_base}}
    """

    QUERY_KPIS_TARJETA = f"""
        SELECT
            SUM(CASE WHEN Respuesta = 'Débito'  THEN 1 ELSE 0 END) AS total_debito,
            SUM(CASE WHEN Respuesta = 'Crédito' THEN 1 ELSE 0 END) AS total_credito
        FROM {VISTA}
        WHERE orden = 3
        {{filtros_base}}
    """

    QUERY_PLANES = f"""
        SELECT Respuesta AS TipoPlan, COUNT(*) AS total
        FROM {VISTA}
        WHERE orden = 1 AND Respuesta IS NOT NULL
        GROUP BY Respuesta
        ORDER BY total DESC
    """

    QUERY_TARJETAS = f"""
        SELECT Respuesta AS TipoTarjeta, COUNT(*) AS total
        FROM {VISTA}
        WHERE orden = 3 AND Respuesta IS NOT NULL
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
        """Con alias v1/v3 para queries con JOIN."""
        condiciones = []
        params = []

        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(v1.Cedula) = %s")
            params.append(filtros["cedula"].strip())

        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(v1.Nombre) = %s")
            params.append(filtros["nombre"].strip())

        if filtros.get("correo"):
            condiciones.append("AND RTRIM(v1.Correo) = %s")
            params.append(filtros["correo"].strip())

        if filtros.get("tipo_plan"):
            condiciones.append("AND v1.Respuesta = %s")
            params.append(filtros["tipo_plan"])

        if filtros.get("tipo_tarjeta"):
            condiciones.append("AND v3.Respuesta = %s")
            params.append(filtros["tipo_tarjeta"])

        if filtros.get("fecha_inicio"):
            condiciones.append("AND CAST(v1.FechaHora AS DATE) >= %s")
            params.append(filtros["fecha_inicio"])

        if filtros.get("fecha_fin"):
            condiciones.append("AND CAST(v1.FechaHora AS DATE) <= %s")
            params.append(filtros["fecha_fin"])

        return " ".join(condiciones), params

    @staticmethod
    def _construir_filtros_simple(filtros: dict):
        """Sin alias, para queries simples (KPIS, planes, tarjetas)."""
        condiciones = []
        params = []

        if filtros.get("cedula"):
            condiciones.append("AND RTRIM(Cedula) = %s")
            params.append(filtros["cedula"].strip())

        if filtros.get("nombre"):
            condiciones.append("AND RTRIM(Nombre) = %s")
            params.append(filtros["nombre"].strip())

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
        sql_filtros, params = cls._construir_filtros_simple(filtros or {})
        sql_plan    = cls.QUERY_KPIS.format(filtros_base=sql_filtros)
        sql_tarjeta = cls.QUERY_KPIS_TARJETA.format(filtros_base=sql_filtros)
        r_plan    = ReportesDBService.ejecutar_query(sql_plan, params)
        r_tarjeta = ReportesDBService.ejecutar_query(sql_tarjeta, params)
        kpis = r_plan[0] if r_plan else {}
        if r_tarjeta:
            kpis.update(r_tarjeta[0])
        return kpis

    @classmethod
    def obtener_planes(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_PLANES)

    @classmethod
    def obtener_tarjetas(cls) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_TARJETAS)

    @classmethod
    def buscar_opciones(cls, columna: str, termino: str) -> list[dict]:
        sql = cls.QUERY_BUSCAR_OPCIONES.format(col=columna, vista=VISTA)
        return ReportesDBService.ejecutar_query(sql, [f"%{termino}%"])