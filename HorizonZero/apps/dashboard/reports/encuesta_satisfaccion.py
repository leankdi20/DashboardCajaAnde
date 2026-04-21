from ..services.db_service import ReportesDBService


class ReporteEncuestaSatisfaccion:

    QUERY = """
        WITH calificaciones AS (
            SELECT *,
                CASE 
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy fácil'           THEN 5
                            WHEN 'Fácil'               THEN 4
                            WHEN 'Ni fácil ni difícil' THEN 3
                            WHEN 'Difícil'             THEN 2
                            WHEN 'Muy difícil'         THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END as valor_numerico
            FROM dbo.vw_reporte_encuestas_satisfaccion
        ),
        promedios AS (
            SELECT respuesta_id, AVG(CAST(valor_numerico AS FLOAT)) as promedio_encuesta
            FROM calificaciones
            WHERE valor_numerico IS NOT NULL
            GROUP BY respuesta_id
        )
        SELECT DISTINCT
            c.respuesta_id, c.encuesta_id, c.Fecha, c.Hora,
            c.Cedula, c.Nombre, c.Correo, c.Agente,
            c.Sucursal, c.Unidad, c.Gestion,
            c.orden, c.encuesta_det_id, c.Pregunta, c.Respuesta,
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


    QUERY_FILTRADO = """
        WITH calificaciones AS (
            SELECT *,
                CASE 
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy fácil'           THEN 5
                            WHEN 'Fácil'               THEN 4
                            WHEN 'Ni fácil ni difícil' THEN 3
                            WHEN 'Difícil'             THEN 2
                            WHEN 'Muy difícil'         THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END as valor_numerico
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE 1=1
            {filtros_base}
        ),
        promedios AS (
            SELECT respuesta_id, AVG(CAST(valor_numerico AS FLOAT)) as promedio_encuesta
            FROM calificaciones
            WHERE valor_numerico IS NOT NULL
            GROUP BY respuesta_id
        )
        SELECT DISTINCT
            c.respuesta_id, c.encuesta_id, c.Fecha, c.Hora,
            c.Cedula, c.Nombre, c.Correo, c.Agente,
            c.Sucursal, c.Unidad, c.Gestion,
            c.orden, c.encuesta_det_id, c.Pregunta, c.Respuesta,
            p.promedio_encuesta,
            CASE 
                WHEN p.promedio_encuesta >= 4 THEN 'promotor'
                WHEN p.promedio_encuesta = 3  THEN 'pasivo'
                WHEN p.promedio_encuesta <= 2 THEN 'detractor'
                ELSE NULL
            END as clasificacion
        FROM calificaciones c
        JOIN promedios p ON c.respuesta_id = p.respuesta_id
        {filtro_clasificacion}
        ORDER BY c.Fecha DESC
    """

    QUERY_DETALLE = """
        SELECT *
        FROM dbo.vw_reporte_encuestas_satisfaccion
        WHERE respuesta_id = %s
        ORDER BY orden
    """


    QUERY_PROMEDIO_AGENTE = """
        SELECT 
            COUNT(DISTINCT respuesta_id) as total_encuestas,
            AVG(CAST(
                CASE 
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy fácil'             THEN 5
                            WHEN 'Fácil'                 THEN 4
                            WHEN 'Ni fácil ni difícil'   THEN 3
                            WHEN 'Difícil'               THEN 2
                            WHEN 'Muy difícil'           THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END AS FLOAT
            )) as promedio_agente
        FROM dbo.vw_reporte_encuestas_satisfaccion
        WHERE Agente = %s
        AND Pregunta IN (
            '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
            '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
            '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
        )
    """

    QUERY_PROMEDIO_ENCUESTA = """
        SELECT 
            AVG(CAST(
                CASE 
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy fácil'             THEN 5
                            WHEN 'Fácil'                 THEN 4
                            WHEN 'Ni fácil ni difícil'   THEN 3
                            WHEN 'Difícil'               THEN 2
                            WHEN 'Muy difícil'           THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END AS FLOAT
            )) as promedio_encuesta
        FROM dbo.vw_reporte_encuestas_satisfaccion
        WHERE respuesta_id = %s
        AND Pregunta IN (
            '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
            '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
            '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
        )
    """
    
    QUERY_KPIS_GLOBALES = """
        WITH calificaciones AS (
            SELECT
                respuesta_id,
                Sucursal,
                Fecha,
                CASE 
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy fácil'           THEN 5
                            WHEN 'Fácil'               THEN 4
                            WHEN 'Ni fácil ni difícil' THEN 3
                            WHEN 'Difícil'             THEN 2
                            WHEN 'Muy difícil'         THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta
                            WHEN 'Muy satisfecho'                THEN 5
                            WHEN 'Satisfecho'                    THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                            WHEN 'Poco satisfecho'               THEN 2
                            WHEN 'Nada satisfecho'               THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END as valor_numerico,
                 -- ── NUEVO: valor específico de dificultad ──────────
                CASE 
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta
                            WHEN 'Muy fácil'           THEN 5
                            WHEN 'Fácil'               THEN 4
                            WHEN 'Ni fácil ni difícil' THEN 3
                            WHEN 'Difícil'             THEN 2
                            WHEN 'Muy difícil'         THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END as valor_dificultad,
                -- ────────────────────────────────────────────────────
                CASE 
                    WHEN Pregunta = '¿Recomienda los productos y servicios de Caja de ANDE a otro accionista?' 
                    AND Respuesta = '1' THEN 1 ELSE 0 
                END as es_promotor_lealtad,
                CASE 
                    WHEN Pregunta = '¿Recomienda los productos y servicios de Caja de ANDE a otro accionista?' 
                    THEN 1 ELSE NULL 
                END as es_pregunta_lealtad
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE Pregunta IN (
                '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
                '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
                '¿Cómo califica su experiencia cuando visita Caja de ANDE?',
                '¿Recomienda los productos y servicios de Caja de ANDE a otro accionista?'
            )
            {filtros}
        )
        SELECT
            COUNT(DISTINCT respuesta_id)                              as total_encuestas,
            AVG(CAST(valor_numerico AS FLOAT))                        as promedio_general,
            AVG(CAST(valor_dificultad AS FLOAT))                      as promedio_dificultad,
            SUM(es_promotor_lealtad)                                  as promotores_lealtad,
            COUNT(es_pregunta_lealtad)                                as total_lealtad,
            SUM(CASE WHEN valor_numerico >= 4 THEN 1 ELSE 0 END)     as promotores_sat,
            SUM(CASE WHEN valor_numerico = 3  THEN 1 ELSE 0 END)     as pasivos,
            SUM(CASE WHEN valor_numerico <= 2 
                    AND valor_numerico IS NOT NULL 
                    THEN 1 ELSE 0 END)                               as detractores,
            COUNT(valor_numerico)                                     as total_satisfaccion
        FROM calificaciones
    """



    @classmethod
    def obtener_kpis_globales(cls, sucursales: list = None, fecha_inicio: str = None, fecha_fin: str = None) -> dict:
        condiciones = []
        params = []

        if sucursales:
            placeholders = ",".join(["%s"] * len(sucursales))
            condiciones.append(f"AND Sucursal IN ({placeholders})")
            params.extend(sucursales)

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
        promedio         = r.get("promedio_general") or 0
        promotores_lealt = r.get("promotores_lealtad") or 0
        total_lealtad    = r.get("total_lealtad") or 1
        promotores_sat   = r.get("promotores_sat") or 0
        pasivos          = r.get("pasivos") or 0
        detractores      = r.get("detractores") or 0
        total_sat        = r.get("total_satisfaccion") or 1
        promedio_dificultad = r.get("promedio_dificultad") or 0

        return {
            "total_encuestas":  r.get("total_encuestas", 0),
            "promedio_general": round((promedio / 5) * 100) if promedio else 0, 
            "indicador_dificultad": round((promedio_dificultad / 5) * 100) if promedio_dificultad else 0,
            "lealtad_pct":      round((promotores_lealt / total_lealtad) * 100) if total_lealtad else 0,
            "promotores":       promotores_sat,
            "promotores_pct":   round((promotores_sat / total_sat) * 100) if total_sat else 0,
            "pasivos":          pasivos,
            "pasivos_pct":      round((pasivos / total_sat) * 100) if total_sat else 0,
            "detractores":      detractores,
            "detractores_pct":  round((detractores / total_sat) * 100) if total_sat else 0,
        }
    


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
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        if not filtros or not any(filtros.values()):
            return ReportesDBService.ejecutar_query(cls.QUERY)

        condiciones_base = []
        params = []

        if filtros.get("agente"):
            condiciones_base.append("AND Agente = %s")
            params.append(filtros["agente"])
        if filtros.get("sucursal"):
            condiciones_base.append("AND Sucursal = %s")
            params.append(filtros["sucursal"])
        if filtros.get("unidad"):
            condiciones_base.append("AND Unidad = %s")
            params.append(filtros["unidad"])
        if filtros.get("gestion"):
            condiciones_base.append("AND Gestion = %s")
            params.append(filtros["gestion"])
        if filtros.get("cedula"):
            condiciones_base.append("AND Cedula = %s")
            params.append(filtros["cedula"])
        if filtros.get("nombre"):
            condiciones_base.append("AND Nombre = %s")
            params.append(filtros["nombre"])
        if filtros.get("fecha_inicio"):
            condiciones_base.append("AND Fecha >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones_base.append("AND Fecha <= %s")
            params.append(filtros["fecha_fin"] + " 23:59:59")

        # Filtro clasificacion va FUERA del CTE sobre el promedio calculado
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
    
