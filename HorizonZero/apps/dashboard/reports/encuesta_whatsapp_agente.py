from ..services.db_service import ReportesDBService

class ReporteEncuestaWhatsappAgente:

    VISTA = "dbo.vw_reporte_encuestas_satisfaccion_whatsapp_agente"

    PREGUNTAS_ESCALA = [
        "¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?",
        "¿La persona que le atendió le brindó respuesta a todas sus consultas?",
        "¿Los tiempos de respuesta fueron adecuados?",
    ]

    QUERY = f"""
        ;WITH calificaciones AS (
            SELECT *,
                CASE
                    WHEN ISNUMERIC(Respuesta) = 1 AND Pregunta IN (
                        '¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?',
                        '¿La persona que le atendió le brindó respuesta a todas sus consultas?',
                        '¿Los tiempos de respuesta fueron adecuados?'
                    ) THEN CAST(Respuesta AS FLOAT)
                    ELSE NULL
                END as valor_numerico
            FROM {VISTA}
        ),
        promedios AS (
            SELECT respuesta_id, AVG(valor_numerico) as promedio_encuesta
            FROM calificaciones
            WHERE valor_numerico IS NOT NULL
            GROUP BY respuesta_id
        )
        SELECT DISTINCT
            c.respuesta_id, c.encuesta_id, c.Fecha, c.Hora,
            c.Cedula, c.Nombre, c.Agente,
            c.encuesta_det_id, c.Pregunta, c.Respuesta,
            p.promedio_encuesta,
            CASE
                WHEN p.promedio_encuesta >= 4 THEN 'promotor'
                WHEN p.promedio_encuesta = 3  THEN 'pasivo'
                WHEN p.promedio_encuesta <= 2 THEN 'detractor'
                ELSE NULL
            END as clasificacion
        FROM calificaciones c
        JOIN promedios p ON c.respuesta_id = p.respuesta_id
        ORDER BY c.Fecha DESC
    """

    QUERY_FILTRADO = f"""
        ;WITH calificaciones AS (
            SELECT *,
                CASE
                    WHEN ISNUMERIC(Respuesta) = 1 AND Pregunta IN (
                        '¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?',
                        '¿La persona que le atendió le brindó respuesta a todas sus consultas?',
                        '¿Los tiempos de respuesta fueron adecuados?'
                    ) THEN CAST(Respuesta AS FLOAT)
                    ELSE NULL
                END as valor_numerico
            FROM {VISTA}
            WHERE 1=1
            {{filtros_base}}
        ),
        promedios AS (
            SELECT respuesta_id, AVG(valor_numerico) as promedio_encuesta
            FROM calificaciones
            WHERE valor_numerico IS NOT NULL
            GROUP BY respuesta_id
        )
        SELECT DISTINCT
            c.respuesta_id, c.encuesta_id, c.Fecha, c.Hora,
            c.Cedula, c.Nombre, c.Agente,
            c.encuesta_det_id, c.Pregunta, c.Respuesta,
            p.promedio_encuesta,
            CASE
                WHEN p.promedio_encuesta >= 4 THEN 'promotor'
                WHEN p.promedio_encuesta = 3  THEN 'pasivo'
                WHEN p.promedio_encuesta <= 2 THEN 'detractor'
                ELSE NULL
            END as clasificacion
        FROM calificaciones c
        JOIN promedios p ON c.respuesta_id = p.respuesta_id
        {{filtro_clasificacion}}
        ORDER BY c.Fecha DESC
    """

    QUERY_DETALLE = f"""
        SELECT * FROM {VISTA}
        WHERE respuesta_id = %s
        ORDER BY encuesta_det_id
    """

    QUERY_PROMEDIO_AGENTE = f"""
        SELECT
            COUNT(DISTINCT respuesta_id) as total_encuestas,
            AVG(CAST(
                CASE WHEN ISNUMERIC(Respuesta) = 1 THEN CAST(Respuesta AS FLOAT) ELSE NULL END
            AS FLOAT)) as promedio_agente
        FROM {VISTA}
        WHERE Agente = %s
        AND Pregunta IN (
            '¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?',
            '¿La persona que le atendió le brindó respuesta a todas sus consultas?',
            '¿Los tiempos de respuesta fueron adecuados?'
        )
    """

    QUERY_PROMEDIO_ENCUESTA = f"""
        SELECT
            AVG(CAST(
                CASE WHEN ISNUMERIC(Respuesta) = 1 THEN CAST(Respuesta AS FLOAT) ELSE NULL END
            AS FLOAT)) as promedio_encuesta
        FROM {VISTA}
        WHERE respuesta_id = %s
        AND Pregunta IN (
            '¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?',
            '¿La persona que le atendió le brindó respuesta a todas sus consultas?',
            '¿Los tiempos de respuesta fueron adecuados?'
        )
    """

    QUERY_KPIS_GLOBALES = f"""
        ;WITH calificaciones AS (
            SELECT
                respuesta_id,
                Agente,
                Fecha,
                CASE WHEN ISNUMERIC(Respuesta) = 1 THEN CAST(Respuesta AS FLOAT) ELSE NULL END as valor_numerico
            FROM {VISTA}
            WHERE Pregunta IN (
                '¿Cuál es su nivel de satisfacción con el servicio brindado por el/la agente?',
                '¿La persona que le atendió le brindó respuesta a todas sus consultas?',
                '¿Los tiempos de respuesta fueron adecuados?'
            )
            {{filtros}}
        )
        SELECT
            COUNT(DISTINCT respuesta_id)                              as total_encuestas,
            AVG(CAST(valor_numerico AS FLOAT))                        as promedio_general,
            SUM(CASE WHEN valor_numerico >= 4 THEN 1 ELSE 0 END)     as promotores_sat,
            SUM(CASE WHEN valor_numerico = 3  THEN 1 ELSE 0 END)     as pasivos,
            SUM(CASE WHEN valor_numerico <= 2
                     AND valor_numerico IS NOT NULL
                     THEN 1 ELSE 0 END)                               as detractores,
            COUNT(valor_numerico)                                     as total_satisfaccion
        FROM calificaciones
    """

    @classmethod
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        if not filtros or not any(filtros.values()):
            return ReportesDBService.ejecutar_query(cls.QUERY)

        condiciones_base = []
        params = []

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
        clasificacion = filtros.get("clasificacion")
        if clasificacion == "promotor":
            filtro_clasificacion = "WHERE p.promedio_encuesta >= 4"
        elif clasificacion == "pasivo":
            filtro_clasificacion = "WHERE p.promedio_encuesta = 3"
        elif clasificacion == "detractor":
            filtro_clasificacion = "WHERE p.promedio_encuesta <= 2"

        sql = cls.QUERY_FILTRADO.format(
            filtros_base=" ".join(condiciones_base),
            filtro_clasificacion=filtro_clasificacion,
        )
        return ReportesDBService.ejecutar_query(sql, params)
    
    @classmethod
    def obtener_datos_agrupados(cls, filtros: dict = None) -> list[dict]:
        datos = cls.obtener_datos(filtros)
        vistos = {}
        for fila in datos:
            rid = fila["respuesta_id"]
            if rid not in vistos:
                vistos[rid] = fila
        return list(vistos.values())

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
    def obtener_kpis_globales(cls, agentes: list = None, fecha_inicio: str = None, fecha_fin: str = None) -> dict:
        condiciones = []
        params = []

        if agentes:
            placeholders = ",".join(["%s"] * len(agentes))
            condiciones.append(f"AND Agente IN ({placeholders})")
            params.extend(agentes)
        if fecha_inicio:
            condiciones.append("AND Fecha >= %s")
            params.append(fecha_inicio)
        if fecha_fin:
            condiciones.append("AND Fecha <= %s")
            params.append(fecha_fin + " 23:59:59")

        sql = cls.QUERY_KPIS_GLOBALES.format(filtros=" ".join(condiciones))
        resultado = ReportesDBService.ejecutar_query(sql, params)

        if not resultado:
            return {}

        r = resultado[0]
        promedio      = r.get("promedio_general") or 0
        promotores    = r.get("promotores_sat") or 0
        pasivos       = r.get("pasivos") or 0
        detractores   = r.get("detractores") or 0
        total_sat     = r.get("total_satisfaccion") or 1

        return {
            "total_encuestas":  r.get("total_encuestas", 0),
            "promedio_general": round((promedio / 5) * 100) if promedio else 0,
            "promotores":       promotores,
            "promotores_pct":   round((promotores / total_sat) * 100) if total_sat else 0,
            "pasivos":          pasivos,
            "pasivos_pct":      round((pasivos / total_sat) * 100) if total_sat else 0,
            "detractores":      detractores,
            "detractores_pct":  round((detractores / total_sat) * 100) if total_sat else 0,
        }



