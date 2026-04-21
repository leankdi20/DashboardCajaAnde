from ..services.db_service import ReportesDBService

class ReporteEncuestaPaginaWeb:

    PREGUNTA_SITIO = "¿Cuál sitio va a calificar?"
    PREGUNTA_NAVEGACION = "¿Considera que el sitio web es amigable y fácil de navegar?"
    PREGUNTA_INFO = "¿Encontró fácilmente la información que buscaba?"
    

    VISTA = "dbo.vw_reporte_encuestas_satisfaccion_pagina_web"

    QUERY = f"SELECT * FROM {VISTA} ORDER BY Fecha DESC"

    QUERY_FILTRADO = f"""
        SELECT * FROM {VISTA}
        WHERE 1=1
        {{filtros}}
        ORDER BY Fecha DESC
    """

    QUERY_DETALLE = """
        SELECT
            respuesta_id,
            encuesta_id,
            encuesta_nombre,
            encuesta_correos,
            encuesta_creado,
            Fecha,
            Nombre,
            Pregunta,
            Respuesta
        FROM dbo.vw_reporte_encuestas_satisfaccion_pagina_web
        WHERE respuesta_id = %s
        ORDER BY Pregunta
    """

    PREGUNTA_NPS = "¿Considera que el sitio web es amigable y fácil de navegar?"

    QUERY_PROMEDIO_ENCUESTA = f"""
        SELECT
            AVG(CAST(
                CASE WHEN ISNUMERIC(Respuesta) = 1 
                    THEN CAST(Respuesta AS FLOAT) 
                    ELSE NULL END AS FLOAT
            )) as promedio_encuesta
        FROM {VISTA}
        WHERE respuesta_id = %s
        AND Pregunta = '¿Considera que el sitio web es amigable y fácil de navegar?'
    """

    QUERY_KPIS_GLOBALES = f"""
        SELECT
            COUNT(*) AS total_encuestas,
            AVG(t.promedio_encuesta) AS promedio_general,
            SUM(CASE WHEN t.promedio_encuesta >= 9 THEN 1 ELSE 0 END) AS promotores,
            SUM(CASE WHEN t.promedio_encuesta >= 7 AND t.promedio_encuesta < 9 THEN 1 ELSE 0 END) AS pasivos,
            SUM(CASE WHEN t.promedio_encuesta < 7 THEN 1 ELSE 0 END) AS detractores,
            COUNT(*) AS total_respuestas
        FROM (
            SELECT
                v.respuesta_id,
                AVG(CAST(v.Respuesta AS FLOAT)) AS promedio_encuesta
            FROM {VISTA} v
            WHERE v.Pregunta = %s
            AND ISNUMERIC(v.Respuesta) = 1
            {{filtros}}
            {{filtro_sitio}}
            GROUP BY v.respuesta_id
        ) t
    """
    QUERY_KPI_INFO = f"""
        SELECT
            COUNT(*) AS total_info,
            SUM(
                CASE
                    WHEN Respuesta IN ('Sí', 'Si', '1', 'SI')
                    THEN 1
                    ELSE 0
                END
            ) AS respuestas_si
        FROM {VISTA}
        WHERE Pregunta = %s
        {{filtros}}
        {{filtro_sitio}}
    """

    @classmethod
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        filtros = filtros or {}

        condiciones = []
        params = []

        if filtros.get("nombre"):
            condiciones.append("AND base.Nombre = %s")
            params.append(filtros["nombre"])

        if filtros.get("fecha_inicio"):
            condiciones.append("AND base.Fecha >= %s")
            params.append(filtros["fecha_inicio"])

        if filtros.get("fecha_fin"):
            condiciones.append("AND base.Fecha <= %s")
            params.append(filtros["fecha_fin"] + " 23:59:59")

        if filtros.get("sitio_evaluado"):
            condiciones.append(f"""
                AND base.respuesta_id IN (
                    SELECT respuesta_id
                    FROM {cls.VISTA}
                    WHERE Pregunta = %s AND Respuesta = %s
                )
            """)
            params.append(cls.PREGUNTA_SITIO)
            params.append(filtros["sitio_evaluado"])

        sql = f"""
            ;WITH promedios AS (
                SELECT
                    respuesta_id,
                    AVG(CAST(Respuesta AS FLOAT)) AS promedio_kpi
                FROM {cls.VISTA}
                WHERE Pregunta = %s
                AND ISNUMERIC(Respuesta) = 1
                GROUP BY respuesta_id
            ),
            segmentos AS (
                SELECT
                    respuesta_id,
                    CASE
                        WHEN promedio_kpi >= 9 THEN 'Promotor'
                        WHEN promedio_kpi >= 7 THEN 'Pasivo'
                        ELSE 'Detractor'
                    END AS segmento_nps
                FROM promedios
            )
            SELECT
                base.*,
                seg.segmento_nps
            FROM {cls.VISTA} base
            LEFT JOIN segmentos seg
                ON seg.respuesta_id = base.respuesta_id
            WHERE 1=1
            {' '.join(condiciones)}
        """

        final_params = [cls.PREGUNTA_NAVEGACION] + params

        if filtros.get("segmento_nps"):
            sql += " AND seg.segmento_nps = %s"
            final_params.append(filtros["segmento_nps"])

        sql += " ORDER BY base.Fecha DESC"

        return ReportesDBService.ejecutar_query(sql, final_params)
    

    @classmethod
    def obtener_datos_agrupados(cls, filtros: dict = None) -> list[dict]:
        datos = cls.obtener_datos(filtros)
        vistos = {}

        for fila in datos:
            rid = fila["respuesta_id"]
            if rid not in vistos:
                vistos[rid] = {
                    "respuesta_id": rid,
                    "Fecha": fila.get("Fecha"),
                    "Nombre": fila.get("Nombre"),
                    "segmento_nps": fila.get("segmento_nps", ""),
                }

        return list(vistos.values())

   
    @classmethod
    def obtener_detalle(cls, respuesta_id: int) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_DETALLE, [respuesta_id])

    @classmethod
    def obtener_kpis_globales(
        cls,
        fecha_inicio: str = None,
        fecha_fin: str = None,
        sitio_evaluado: str = None,
    ) -> dict:
        condiciones = []
        filtro_sitio = []
        params_nps = [cls.PREGUNTA_NAVEGACION]
        params_info = [cls.PREGUNTA_INFO]

        if fecha_inicio:
            condiciones.append("AND Fecha >= %s")
            params_nps.append(fecha_inicio)
            params_info.append(fecha_inicio)

        if fecha_fin:
            condiciones.append("AND Fecha <= %s")
            params_nps.append(fecha_fin + " 23:59:59")
            params_info.append(fecha_fin + " 23:59:59")

        if sitio_evaluado:
            filtro_sitio.append(f"""
                AND respuesta_id IN (
                    SELECT s.respuesta_id
                    FROM {cls.VISTA} s
                    WHERE s.Pregunta = %s
                    AND s.Respuesta = %s
                )
            """)
            params_nps.extend([cls.PREGUNTA_SITIO, sitio_evaluado])
            params_info.extend([cls.PREGUNTA_SITIO, sitio_evaluado])

        sql_nps = cls.QUERY_KPIS_GLOBALES.format(
            filtros=" ".join(condiciones),
            filtro_sitio=" ".join(filtro_sitio),
        )

        sql_info = cls.QUERY_KPI_INFO.format(
            filtros=" ".join(condiciones),
            filtro_sitio=" ".join(filtro_sitio),
        )

        resultado_nps = ReportesDBService.ejecutar_query(sql_nps, params_nps)
        resultado_info = ReportesDBService.ejecutar_query(sql_info, params_info)

        if not resultado_nps:
            return {
                "total_encuestas": 0,
                "promedio_general": 0,
                "nps": 0,
                "promotores_pct": 0,
                "pasivos_pct": 0,
                "detractores_pct": 0,
                "encontro_info_pct": 0,
                "promotores": 0,
                "pasivos": 0,
                "detractores": 0,
                "total_respuestas": 0,
            }

        r = resultado_nps[0]
        promedio = r.get("promedio_general") or 0
        promotores = r.get("promotores") or 0
        pasivos = r.get("pasivos") or 0
        detractores = r.get("detractores") or 0
        total = r.get("total_respuestas") or 0

        info = resultado_info[0] if resultado_info else {}
        total_info = info.get("total_info") or 0
        respuestas_si = info.get("respuestas_si") or 0
        encontro_info_pct = round((respuestas_si / total_info) * 100) if total_info else 0

        if total == 0:
            return {
                "total_encuestas": 0,
                "promedio_general": 0,
                "nps": 0,
                "promotores_pct": 0,
                "pasivos_pct": 0,
                "detractores_pct": 0,
                "encontro_info_pct": encontro_info_pct,
                "promotores": 0,
                "pasivos": 0,
                "detractores": 0,
                "total_respuestas": 0,
            }

        nps = round(((promotores - detractores) / total) * 100)

        return {
            "total_encuestas": r.get("total_encuestas", 0),
            "promedio_general": round(promedio, 1),
            "nps": nps,
            "promotores_pct": round((promotores / total) * 100),
            "pasivos_pct": round((pasivos / total) * 100),
            "detractores_pct": round((detractores / total) * 100),
            "encontro_info_pct": encontro_info_pct,
            "promotores": promotores,
            "pasivos": pasivos,
            "detractores": detractores,
            "total_respuestas": total,
        }
    
    @classmethod
    def obtener_promedio_encuesta(cls, respuesta_id: int) -> float:
        resultado = ReportesDBService.ejecutar_query(cls.QUERY_PROMEDIO_ENCUESTA, [respuesta_id])
        if resultado:
            return round(resultado[0].get("promedio_encuesta") or 0, 2)
        return 0


    @classmethod
    def obtener_opciones_sitio(cls) -> list[dict]:
        sql = f"""
            SELECT DISTINCT Respuesta AS sitio
            FROM {cls.VISTA}
            WHERE Pregunta = %s
            AND Respuesta IS NOT NULL
            AND LTRIM(RTRIM(Respuesta)) <> ''
            ORDER BY Respuesta
        """
        return ReportesDBService.ejecutar_query(sql, [cls.PREGUNTA_SITIO])