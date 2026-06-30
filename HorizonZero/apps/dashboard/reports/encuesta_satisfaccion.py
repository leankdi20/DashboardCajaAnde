from ..services.db_service import ReportesDBService
from datetime import datetime, timedelta
import time

class ReporteEncuestaSatisfaccion:

    # ════════════════════════════════════════════════════════════════
    # QUERY PRINCIPAL
    #
    # La vista se lee UNA SOLA VEZ en el CTE `base`.
    # Todos los CTEs siguientes operan sobre `base`, nunca
    # vuelven a tocar dbo.vw_reporte_encuestas_satisfaccion.
    #
    # {filtros_base}          → condiciones WHERE opcionales (AND ...)
    # {filtro_clasificacion}  → WHERE sobre promedio calculado
    # ════════════════════════════════════════════════════════════════

    QUERY_DATOS = """
        WITH base AS (
            SELECT *
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE 1=1
            {filtros_base}
        ),
        calificaciones AS (
            SELECT
                respuesta_id,
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
                END AS valor_numerico
            FROM base
        ),
        promedios AS (
            SELECT
                respuesta_id,
                AVG(CAST(valor_numerico AS FLOAT)) AS promedio_encuesta
            FROM calificaciones
            WHERE valor_numerico IS NOT NULL
            GROUP BY respuesta_id
        ),
        metadatos AS (
            -- Cabecera de cada encuesta — una fila por respuesta_id
            SELECT
                respuesta_id,
                MAX(encuesta_id) AS encuesta_id,
                MAX(Fecha)       AS Fecha,
                MAX(Hora)        AS Hora,
                MAX(Cedula)      AS Cedula,
                MAX(Nombre)      AS Nombre,
                MAX(Correo)      AS Correo,
                MAX(Agente)      AS Agente,
                MAX(Sucursal)    AS Sucursal,
                MAX(Unidad)      AS Unidad,
                MAX(Gestion)     AS Gestion
            FROM base
            GROUP BY respuesta_id
        )
        SELECT
            m.respuesta_id, m.encuesta_id, m.Fecha, m.Hora,
            m.Cedula, m.Nombre, m.Correo, m.Agente,
            m.Sucursal, m.Unidad, m.Gestion,
            p.promedio_encuesta,
            CASE
                WHEN p.promedio_encuesta >= 4 THEN 'promotor'
                WHEN p.promedio_encuesta = 3  THEN 'pasivo'
                WHEN p.promedio_encuesta <= 2 THEN 'detractor'
                ELSE NULL
            END AS clasificacion
        FROM metadatos m
        JOIN promedios p ON m.respuesta_id = p.respuesta_id
        {filtro_clasificacion}
        ORDER BY m.Fecha DESC
    """

    # ════════════════════════════════════════════════════════════════
    # KPIs GLOBALES — también lee la vista una sola vez
    # ════════════════════════════════════════════════════════════════

    QUERY_KPIS_GLOBALES = """
        WITH base AS (
            SELECT *
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE Pregunta IN (
                '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
                '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
                '¿Cómo califica su experiencia cuando visita Caja de ANDE?',
                '¿Recomienda los productos y servicios de Caja de ANDE a otro accionista?'
            )
            {filtros}
        ),
        calificaciones AS (
            SELECT
                respuesta_id,
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
                END AS valor_numerico,
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
                END AS valor_dificultad,
                CASE
                    WHEN Pregunta = '¿Recomienda los productos y servicios de Caja de ANDE a otro accionista?'
                         AND Respuesta = '1' THEN 1
                    ELSE 0
                END AS es_promotor_lealtad,
                CASE
                    WHEN Pregunta = '¿Recomienda los productos y servicios de Caja de ANDE a otro accionista?'
                    THEN 1 ELSE NULL
                END AS es_pregunta_lealtad
            FROM base
        )
        SELECT
            COUNT(DISTINCT respuesta_id)                             AS total_encuestas,
            AVG(CAST(valor_numerico   AS FLOAT))                     AS promedio_general,
            AVG(CAST(valor_dificultad AS FLOAT))                     AS promedio_dificultad,
            SUM(es_promotor_lealtad)                                 AS promotores_lealtad,
            COUNT(es_pregunta_lealtad)                               AS total_lealtad,
            SUM(CASE WHEN valor_numerico >= 4 THEN 1 ELSE 0 END)    AS promotores_sat,
            SUM(CASE WHEN valor_numerico = 3  THEN 1 ELSE 0 END)    AS pasivos,
            SUM(CASE WHEN valor_numerico <= 2
                     AND valor_numerico IS NOT NULL
                     THEN 1 ELSE 0 END)                              AS detractores,
            COUNT(valor_numerico)                                    AS total_satisfaccion
        FROM calificaciones
    """

    # ════════════════════════════════════════════════════════════════
    # TIMELINE — agrupación ligera, sin traer todas las columnas
    # ════════════════════════════════════════════════════════════════

    QUERY_TIMELINE = """
        SELECT
            YEAR(Fecha)                    AS anio,
            MONTH(Fecha)                   AS mes,
            COUNT(DISTINCT respuesta_id)   AS total
        FROM dbo.vw_reporte_encuestas_satisfaccion
        WHERE Fecha IS NOT NULL
        GROUP BY YEAR(Fecha), MONTH(Fecha)
        ORDER BY anio ASC, mes ASC
    """

    # ════════════════════════════════════════════════════════════════
    # DETALLE — todas las preguntas de una encuesta específica
    # ════════════════════════════════════════════════════════════════

    QUERY_DETALLE = """
        SELECT *
        FROM dbo.vw_reporte_encuestas_satisfaccion
        WHERE respuesta_id = %s
        ORDER BY orden
    """

    # ════════════════════════════════════════════════════════════════
    # PROMEDIO AGENTE y PROMEDIO ENCUESTA — queries de detalle
    # ════════════════════════════════════════════════════════════════

    QUERY_PROMEDIO_AGENTE = """
        WITH base AS (
            SELECT respuesta_id, Respuesta, Pregunta
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE Agente = %s
            AND Pregunta IN (
                '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
                '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
                '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
            )
        )
        SELECT
            COUNT(DISTINCT respuesta_id) AS total_encuestas,
            AVG(CAST(
                CASE
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta WHEN 'Muy satisfecho' THEN 5 WHEN 'Satisfecho' THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3 WHEN 'Poco satisfecho' THEN 2
                            WHEN 'Nada satisfecho' THEN 1 ELSE NULL END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta WHEN 'Muy fácil' THEN 5 WHEN 'Fácil' THEN 4
                            WHEN 'Ni fácil ni difícil' THEN 3 WHEN 'Difícil' THEN 2
                            WHEN 'Muy difícil' THEN 1 ELSE NULL END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta WHEN 'Muy satisfecho' THEN 5 WHEN 'Satisfecho' THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3 WHEN 'Poco satisfecho' THEN 2
                            WHEN 'Nada satisfecho' THEN 1 ELSE NULL END
                    ELSE NULL
                END AS FLOAT
            )) AS promedio_agente
        FROM base
    """
    QUERY_DATOS_CON_PREGUNTAS = """
        WITH base AS (
            SELECT *
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE 1=1
            {filtros_base}
        ),
        promedios AS (
            SELECT respuesta_id,
                AVG(CAST(
                    CASE
                        WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                            CASE Respuesta WHEN 'Muy satisfecho' THEN 5 WHEN 'Satisfecho' THEN 4
                                WHEN 'Ni satisfecho ni insatisfecho' THEN 3 WHEN 'Poco satisfecho' THEN 2
                                WHEN 'Nada satisfecho' THEN 1 ELSE NULL END
                        WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                            CASE Respuesta WHEN 'Muy fácil' THEN 5 WHEN 'Fácil' THEN 4
                                WHEN 'Ni fácil ni difícil' THEN 3 WHEN 'Difícil' THEN 2
                                WHEN 'Muy difícil' THEN 1 ELSE NULL END
                        WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                            CASE Respuesta WHEN 'Muy satisfecho' THEN 5 WHEN 'Satisfecho' THEN 4
                                WHEN 'Ni satisfecho ni insatisfecho' THEN 3 WHEN 'Poco satisfecho' THEN 2
                                WHEN 'Nada satisfecho' THEN 1 ELSE NULL END
                        ELSE NULL
                    END AS FLOAT
                )) AS promedio_encuesta
            FROM base
            GROUP BY respuesta_id
        )
        SELECT
            b.respuesta_id, b.Fecha, b.Hora, b.Cedula, b.Nombre,
            b.Agente, b.Sucursal, b.Unidad, b.Gestion,
            b.Pregunta, b.Respuesta, b.orden,
            p.promedio_encuesta,
            CASE
                WHEN p.promedio_encuesta >= 4 THEN 'promotor'
                WHEN p.promedio_encuesta = 3  THEN 'pasivo'
                WHEN p.promedio_encuesta <= 2 THEN 'detractor'
                ELSE NULL
            END AS clasificacion
        FROM base b
        JOIN promedios p ON b.respuesta_id = p.respuesta_id
        {filtro_clasificacion}
        ORDER BY b.Fecha DESC, b.respuesta_id, b.orden
    """

    QUERY_PROMEDIO_ENCUESTA = """
        WITH base AS (
            SELECT Respuesta, Pregunta
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE respuesta_id = %s
            AND Pregunta IN (
                '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
                '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
                '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
            )
        )
        SELECT
            AVG(CAST(
                CASE
                    WHEN Pregunta = '¿Qué tan satisfecho está con la atención recibida el día de hoy?' THEN
                        CASE Respuesta WHEN 'Muy satisfecho' THEN 5 WHEN 'Satisfecho' THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3 WHEN 'Poco satisfecho' THEN 2
                            WHEN 'Nada satisfecho' THEN 1 ELSE NULL END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar su gestión el día de hoy?' THEN
                        CASE Respuesta WHEN 'Muy fácil' THEN 5 WHEN 'Fácil' THEN 4
                            WHEN 'Ni fácil ni difícil' THEN 3 WHEN 'Difícil' THEN 2
                            WHEN 'Muy difícil' THEN 1 ELSE NULL END
                    WHEN Pregunta = '¿Cómo califica su experiencia cuando visita Caja de ANDE?' THEN
                        CASE Respuesta WHEN 'Muy satisfecho' THEN 5 WHEN 'Satisfecho' THEN 4
                            WHEN 'Ni satisfecho ni insatisfecho' THEN 3 WHEN 'Poco satisfecho' THEN 2
                            WHEN 'Nada satisfecho' THEN 1 ELSE NULL END
                    ELSE NULL
                END AS FLOAT
            )) AS promedio_encuesta
        FROM base
    """

    # ════════════════════════════════════════════════════════════════
    # MÉTODOS
    # ════════════════════════════════════════════════════════════════

    @classmethod
    def obtener_timeline(cls, fecha_inicio: str = None, fecha_fin: str = None) -> list:
        condiciones = []
        params = []

        if fecha_inicio:
            condiciones.append("AND Fecha >= %s")
            params.append(fecha_inicio)

        if fecha_fin:
            condiciones.append("AND Fecha <= %s")
            params.append(fecha_fin + " 23:59:59")

        sql = """
            SELECT
                YEAR(Fecha)                    AS anio,
                MONTH(Fecha)                   AS mes,
                COUNT(DISTINCT respuesta_id)   AS total
            FROM dbo.vw_reporte_encuestas_satisfaccion
            WHERE Fecha IS NOT NULL
            {filtros}
            GROUP BY YEAR(Fecha), MONTH(Fecha)
            ORDER BY anio ASC, mes ASC
        """.format(filtros=" ".join(condiciones))

        return ReportesDBService.ejecutar_query(sql, params)

    @classmethod
    def obtener_kpis_globales(
        cls,
        sucursales: list = None,
        fecha_inicio: str = None,
        fecha_fin: str = None,
    ) -> dict:
        condiciones, params = [], []

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

        sql      = cls.QUERY_KPIS_GLOBALES.format(filtros=" ".join(condiciones))
        resultado = ReportesDBService.ejecutar_query(sql, params)

        if not resultado:
            return {}

        r                   = resultado[0]
        promedio            = r.get("promedio_general")    or 0
        promedio_dificultad = r.get("promedio_dificultad") or 0
        promotores_lealt    = r.get("promotores_lealtad")  or 0
        total_lealtad       = r.get("total_lealtad")       or 1
        promotores_sat      = r.get("promotores_sat")      or 0
        pasivos             = r.get("pasivos")              or 0
        detractores         = r.get("detractores")          or 0
        total_sat           = r.get("total_satisfaccion")  or 1

        return {
            "total_encuestas":      r.get("total_encuestas", 0),
            "promedio_general":     round((promedio            / 5) * 100) if promedio            else 0,
            "indicador_dificultad": round((promedio_dificultad / 5) * 100) if promedio_dificultad else 0,
            "lealtad_pct":          round((promotores_lealt / total_lealtad) * 100) if total_lealtad else 0,
            "promotores":           promotores_sat,
            "promotores_pct":       round((promotores_sat / total_sat) * 100) if total_sat else 0,
            "pasivos":              pasivos,
            "pasivos_pct":          round((pasivos      / total_sat) * 100) if total_sat else 0,
            "detractores":          detractores,
            "detractores_pct":      round((detractores  / total_sat) * 100) if total_sat else 0,
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
        condiciones, params = [], []
        filtros = filtros or {}

        # Filtro por defecto: últimos 30 días si no viene ninguna fecha
        if not filtros.get("fecha_inicio") and not filtros.get("fecha_fin"):
            filtros["fecha_inicio"] = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
            filtros["fecha_fin"] = datetime.now().strftime("%Y-%m-%d")

        if filtros.get("agente"):
            condiciones.append("AND Agente = %s")
            params.append(filtros["agente"])
        if filtros.get("sucursal"):
            condiciones.append("AND Sucursal = %s")
            params.append(filtros["sucursal"])
        if filtros.get("unidad"):
            condiciones.append("AND Unidad = %s")
            params.append(filtros["unidad"])
        if filtros.get("gestion"):
            condiciones.append("AND Gestion = %s")
            params.append(filtros["gestion"])
        if filtros.get("cedula"):
            condiciones.append("AND Cedula = %s")
            params.append(filtros["cedula"])
        if filtros.get("nombre"):
            condiciones.append("AND Nombre = %s")
            params.append(filtros["nombre"])
        if filtros.get("fecha_inicio"):
            condiciones.append("AND Fecha >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones.append("AND Fecha <= %s")
            params.append(filtros["fecha_fin"] + " 23:59:59")

        filtro_clasificacion = ""
        c = filtros.get("clasificacion")
        if c == "promotor":
            filtro_clasificacion = "WHERE p.promedio_encuesta >= 4"
        elif c == "pasivo":
            filtro_clasificacion = "WHERE p.promedio_encuesta = 3"
        elif c == "detractor":
            filtro_clasificacion = "WHERE p.promedio_encuesta <= 2"

        sql = cls.QUERY_DATOS.format(
            filtros_base=" ".join(condiciones),
            filtro_clasificacion=filtro_clasificacion,
        )
        inicio_q = time.time()
        resultado = ReportesDBService.ejecutar_query(sql, params)
       
        return resultado
        # return ReportesDBService.ejecutar_query(sql, params)
    
    @classmethod
    def obtener_datos_exportar(cls, filtros: dict = None) -> list[dict]:
        """Retorna todas las filas con preguntas/respuestas para exportación."""
        condiciones, params = [], []
        if filtros:
            if filtros.get("agente"):
                condiciones.append("AND Agente = %s"); params.append(filtros["agente"])
            if filtros.get("sucursal"):
                condiciones.append("AND Sucursal = %s"); params.append(filtros["sucursal"])
            if filtros.get("unidad"):
                condiciones.append("AND Unidad = %s"); params.append(filtros["unidad"])
            if filtros.get("gestion"):
                condiciones.append("AND Gestion = %s"); params.append(filtros["gestion"])
            if filtros.get("cedula"):
                condiciones.append("AND Cedula = %s"); params.append(filtros["cedula"])
            if filtros.get("nombre"):
                condiciones.append("AND Nombre = %s"); params.append(filtros["nombre"])
            if filtros.get("fecha_inicio"):
                condiciones.append("AND Fecha >= %s"); params.append(filtros["fecha_inicio"])
            if filtros.get("fecha_fin"):
                condiciones.append("AND Fecha <= %s"); params.append(filtros["fecha_fin"] + " 23:59:59")

        filtro_clasificacion = ""
        if filtros:
            c = filtros.get("clasificacion")
            if c == "promotor":   filtro_clasificacion = "WHERE p.promedio_encuesta >= 4"
            elif c == "pasivo":   filtro_clasificacion = "WHERE p.promedio_encuesta = 3"
            elif c == "detractor": filtro_clasificacion = "WHERE p.promedio_encuesta <= 2"

        sql = cls.QUERY_DATOS_CON_PREGUNTAS.format(
            filtros_base=" ".join(condiciones),
            filtro_clasificacion=filtro_clasificacion,
        )
        return ReportesDBService.ejecutar_query(sql, params)

    @classmethod
    def obtener_datos_agrupados(cls, filtros: dict = None) -> list[dict]:
        """Alias mantenido por compatibilidad. La query ya retorna una fila por encuesta."""
        return cls.obtener_datos(filtros)

    @classmethod
    def obtener_detalle(cls, respuesta_id: int) -> list[dict]:
        return ReportesDBService.ejecutar_query(cls.QUERY_DETALLE, [respuesta_id])