from ..services.db_service import ReportesDBService


class ReporteEncuestaOficinaDigital:

    VISTA = "dbo.vw_reporte_encuestas_satisfaccion_oficina_digital"

    # ── Valor numérico por respuesta ─────────────────────────────
    _CASE_RESPUESTA = """
        CASE Respuesta
            WHEN 'Excelente' THEN 3
            WHEN 'Regular'   THEN 2
            WHEN 'Malo'      THEN 1
            ELSE NULL
        END
    """

    QUERY = f"SELECT * FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital ORDER BY Fecha DESC"

    QUERY_FILTRADO = """
        WITH base AS (
            SELECT *
            FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital
            WHERE 1=1
            {filtros_base}
        ),
        promedios AS (
            SELECT
                respuesta_id,
                AVG(CAST(
                    CASE Respuesta
                        WHEN 'Excelente' THEN 3
                        WHEN 'Regular'   THEN 2
                        WHEN 'Malo'      THEN 1
                        ELSE NULL
                    END AS FLOAT
                )) AS promedio_enc
            FROM base
            WHERE Respuesta IN ('Excelente','Regular','Malo')
            GROUP BY respuesta_id
        ),
        metadatos AS (
            SELECT
                respuesta_id,
                MAX(encuesta_id) AS encuesta_id,
                MAX(Fecha)       AS Fecha,
                
                MAX(Cedula)      AS Cedula,
                MAX(Nombre)      AS Nombre,
                MAX(Agente)      AS Agente
            FROM base
            GROUP BY respuesta_id
        )
        SELECT
            m.respuesta_id, m.encuesta_id, m.Fecha, 
            m.Cedula, m.Nombre, m.Agente,
            p.promedio_enc,
            CASE
                WHEN p.promedio_enc >= 2.5 THEN 'promotor'
                WHEN p.promedio_enc >= 2   THEN 'pasivo'
                WHEN p.promedio_enc <  2   THEN 'detractor'
                ELSE NULL
            END AS clasificacion
        FROM metadatos m
        JOIN promedios p ON m.respuesta_id = p.respuesta_id
        {filtro_clasificacion}
        ORDER BY m.Fecha DESC
    """

    QUERY_DETALLE = f"""
        SELECT * FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital
        WHERE respuesta_id = %s
        ORDER BY orden
    """

    QUERY_PROMEDIO_AGENTE = f"""
        SELECT
            COUNT(DISTINCT respuesta_id) AS total_encuestas,
            AVG(CAST(
                CASE
                    WHEN Pregunta IN (
                        '¿Cómo valora la atención que le brindó el ejecutivo de servicio?',
                        '¿Cómo valora el espacio físico asignado para la oficina digital?',
                        '¿Cómo califica la velocidad de internet para realizar sus trámites en la Oficina Digital?',
                        '¿Cómo fue su experiencia al realizar sus gestiones en la Oficina Digital?'
                    ) THEN
                        CASE Respuesta
                            WHEN 'Excelente' THEN 3
                            WHEN 'Regular'   THEN 2
                            WHEN 'Malo'      THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END AS FLOAT
            )) AS promedio_agente
        FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital
        WHERE Agente = %s
    """

    QUERY_PROMEDIO_ENCUESTA = f"""
        SELECT
            AVG(CAST(
                CASE Respuesta
                    WHEN 'Excelente' THEN 3
                    WHEN 'Regular'   THEN 2
                    WHEN 'Malo'      THEN 1
                    ELSE NULL
                END AS FLOAT
            )) AS promedio_encuesta
        FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital
        WHERE respuesta_id = %s
    """

    QUERY_TIMELINE = """
        SELECT
            YEAR(Fecha)                  AS anio,
            MONTH(Fecha)                 AS mes,
            COUNT(DISTINCT respuesta_id) AS total
        FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital
        WHERE Fecha IS NOT NULL
        GROUP BY YEAR(Fecha), MONTH(Fecha)
        ORDER BY anio ASC, mes ASC
    """

    # ── KPIs globales — promotores/pasivos/detractores por promedio de encuesta ──
    QUERY_KPIS_GLOBALES = """
        WITH base AS (
            SELECT *
            FROM dbo.vw_reporte_encuestas_satisfaccion_oficina_digital
            WHERE Respuesta IN ('Excelente', 'Regular', 'Malo')
            {filtros}
        ),
        promedios AS (
            SELECT
                respuesta_id,
                AVG(CAST(
                    CASE Respuesta
                        WHEN 'Excelente' THEN 3
                        WHEN 'Regular'   THEN 2
                        WHEN 'Malo'      THEN 1
                        ELSE NULL
                    END AS FLOAT
                )) AS promedio_enc
            FROM base
            GROUP BY respuesta_id
        )
        SELECT
            COUNT(*)                                                       AS total_encuestas,
            AVG(CAST(promedio_enc AS FLOAT))                               AS promedio_general,
            SUM(CASE WHEN promedio_enc >= 2.5 THEN 1 ELSE 0 END)          AS promotores,
            SUM(CASE WHEN promedio_enc >= 2 AND promedio_enc < 2.5
                     THEN 1 ELSE 0 END)                                    AS pasivos,
            SUM(CASE WHEN promedio_enc < 2  THEN 1 ELSE 0 END)            AS detractores,
            COUNT(*)                                                       AS total_clasificados
        FROM promedios
    """

    # ─────────────────────────────────────────────────────────────
    # MÉTODOS
    # ─────────────────────────────────────────────────────────────

    @classmethod
    def obtener_timeline(cls) -> list:
        return ReportesDBService.ejecutar_query(cls.QUERY_TIMELINE)

    @classmethod
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        condiciones_base, params = [], []

        if filtros:
            if filtros.get("agente"):
                condiciones_base.append("AND Agente = %s")
                params.append(filtros["agente"])
            if filtros.get("nombre"):
                condiciones_base.append("AND Nombre = %s")
                params.append(filtros["nombre"])
            if filtros.get("cedula"):
                condiciones_base.append("AND Cedula = %s")
                params.append(filtros["cedula"])
            if filtros.get("fecha_inicio"):
                condiciones_base.append("AND Fecha >= %s")
                params.append(filtros["fecha_inicio"])
            if filtros.get("fecha_fin"):
                condiciones_base.append("AND Fecha <= %s")
                params.append(filtros["fecha_fin"] + " 23:59:59")

        filtro_clasificacion = ""
        if filtros:
            c = filtros.get("clasificacion")
            if c == "promotor":
                filtro_clasificacion = "WHERE p.promedio_enc >= 2.5"
            elif c == "pasivo":
                filtro_clasificacion = "WHERE p.promedio_enc >= 2 AND p.promedio_enc < 2.5"
            elif c == "detractor":
                filtro_clasificacion = "WHERE p.promedio_enc < 2"

        sql = cls.QUERY_FILTRADO.format(
            filtros_base=" ".join(condiciones_base),
            filtro_clasificacion=filtro_clasificacion,
        )
        return ReportesDBService.ejecutar_query(sql, params)

    @classmethod
    def obtener_datos_agrupados(cls, filtros: dict = None) -> list[dict]:
        # La query ya retorna una fila por encuesta via CTE
        return cls.obtener_datos(filtros)

    @classmethod
    def obtener_detalle(cls, respuesta_id: int) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_DETALLE, [respuesta_id])

    @classmethod
    def obtener_promedio_agente(cls, agente: str) -> dict:
        resultado = ReportesDBService.ejecutar_query(cls.QUERY_PROMEDIO_AGENTE, [agente])
        if resultado:
            return {
                "total_encuestas": resultado[0].get("total_encuestas", 0),
                "promedio_agente": round(resultado[0].get("promedio_agente") or 0, 2),
            }
        return {"total_encuestas": 0, "promedio_agente": 0}

    @classmethod
    def obtener_promedio_encuesta(cls, respuesta_id: int) -> float:
        resultado = ReportesDBService.ejecutar_query(cls.QUERY_PROMEDIO_ENCUESTA, [respuesta_id])
        if resultado:
            return round(resultado[0].get("promedio_encuesta") or 0, 2)
        return 0

    @classmethod
    def obtener_kpis_globales(
        cls,
        sucursales: list = None,
        fecha_inicio: str = None,
        fecha_fin: str = None,
    ) -> dict:
        condiciones, params = [], []

        if fecha_inicio:
            condiciones.append("AND Fecha >= %s")
            params.append(fecha_inicio)
        if fecha_fin:
            condiciones.append("AND Fecha <= %s")
            params.append(fecha_fin + " 23:59:59")

        sql      = cls.QUERY_KPIS_GLOBALES.format(filtros=" ".join(condiciones))
        resultado = ReportesDBService.ejecutar_query(sql, params)

        if not resultado:
            return {}

        r              = resultado[0]
        promedio       = r.get("promedio_general") or 0
        promotores     = r.get("promotores")       or 0
        pasivos        = r.get("pasivos")           or 0
        detractores    = r.get("detractores")       or 0
        total          = r.get("total_clasificados") or 1

        return {
            "total_encuestas":  r.get("total_encuestas", 0),
            "promedio_general": round((promedio / 3) * 100) if promedio else 0,
            "promotores":       promotores,
            "promotores_pct":   round((promotores  / total) * 100) if total else 0,
            "pasivos":          pasivos,
            "pasivos_pct":      round((pasivos     / total) * 100) if total else 0,
            "detractores":      detractores,
            "detractores_pct":  round((detractores / total) * 100) if total else 0,
        }