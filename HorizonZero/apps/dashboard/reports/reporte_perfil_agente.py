# ═══════════════════════════════════════════════════════════════════
# apps/dashboard/reports/reporte_perfil_agente.py
# ═══════════════════════════════════════════════════════════════════

from ..services.db_service import ReportesDBService

# ── Vistas SQL centralizadas ─────────────────────────────────────
# Si cambia el nombre de una vista, solo se modifica acá.
VISTA_SAT = "dbo.vw_reporte_encuestas_satisfaccion"
VISTA_OD  = "dbo.vw_reporte_encuestas_satisfaccion_oficina_digital"

# ── Lógica de valoración de satisfacción (escala 1-5) ────────────
VALOR_SAT = """
    CASE
        WHEN Pregunta IN (
            '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
            '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
            '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
        ) THEN
            CASE Respuesta
                WHEN 'Muy satisfecho'                THEN 5
                WHEN 'Satisfecho'                    THEN 4
                WHEN 'Ni satisfecho ni insatisfecho' THEN 3
                WHEN 'Poco satisfecho'               THEN 2
                WHEN 'Nada satisfecho'               THEN 1
                WHEN 'Muy fácil'                     THEN 5
                WHEN 'Fácil'                         THEN 4
                WHEN 'Ni fácil ni difícil'           THEN 3
                WHEN 'Difícil'                       THEN 2
                WHEN 'Muy difícil'                   THEN 1
                ELSE NULL
            END
        ELSE NULL
    END
"""

# ── Preguntas de satisfacción que se evalúan ─────────────────────
PREGUNTAS_SAT = """(
    '¿Qué tan satisfecho está con la atención recibida el día de hoy?',
    '¿Cómo fue su experiencia al realizar su gestión el día de hoy?',
    '¿Cómo califica su experiencia cuando visita Caja de ANDE?'
)"""

# ── Lógica de valoración de Oficina Digital (escala 1-3) ─────────
VALOR_OD = """
    CASE Respuesta
        WHEN 'Excelente' THEN 3
        WHEN 'Regular'   THEN 2
        WHEN 'Malo'      THEN 1
        ELSE NULL
    END
"""

# ── Nombres de meses en español ───────────────────────────────────
MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


class ReportePerfilAgente:

    # ────────────────────────────────────────────────────────────
    # INFO RÁPIDA DEL AGENTE — unidad y sucursal
    # Query liviano — no calcula métricas
    # ────────────────────────────────────────────────────────────
    QUERY_INFO = f"""
        SELECT TOP 1 Agente, Unidad, Sucursal
        FROM {VISTA_SAT}
        WHERE Agente = %s
    """

    # ────────────────────────────────────────────────────────────
    # LISTA DE AGENTES — índice con métricas resumidas
    # ────────────────────────────────────────────────────────────
    QUERY_LISTA = f"""
        WITH cal AS (
            SELECT Agente, Unidad, Sucursal, respuesta_id,
                   {VALOR_SAT} AS v
            FROM {VISTA_SAT}
            WHERE Pregunta IN {PREGUNTAS_SAT}
            {{filtros}}
        ),
        prom AS (
            SELECT Agente, Unidad, Sucursal, respuesta_id,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY Agente, Unidad, Sucursal, respuesta_id
        )
        SELECT
            Agente, Unidad, Sucursal,
            COUNT(DISTINCT respuesta_id)                            AS total_enc,
            ROUND(AVG(p), 2)                                        AS promedio,
            SUM(CASE WHEN p >= 4 THEN 1 ELSE 0 END)                AS promotores,
            SUM(CASE WHEN p  = 3 THEN 1 ELSE 0 END)                AS pasivos,
            SUM(CASE WHEN p <= 2 THEN 1 ELSE 0 END)                AS detractores,
            RANK() OVER (PARTITION BY Unidad ORDER BY AVG(p) DESC) AS rank_unidad
        FROM prom
        GROUP BY Agente, Unidad, Sucursal
        ORDER BY promedio DESC
    """

    # ────────────────────────────────────────────────────────────
    # KPIs SATISFACCIÓN del agente (escala 1-5)
    # ────────────────────────────────────────────────────────────
    QUERY_KPI_SAT = f"""
        WITH cal AS (
            SELECT respuesta_id, Fecha,
                   {VALOR_SAT} AS v
            FROM {VISTA_SAT}
            WHERE Agente = %s
            AND Pregunta IN {PREGUNTAS_SAT}
            {{filtro_fecha}}
        ),
        prom AS (
            SELECT respuesta_id, Fecha,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, Fecha
        )
        SELECT
            COUNT(DISTINCT respuesta_id)             AS total_enc,
            ROUND(AVG(p), 2)                         AS promedio,
            SUM(CASE WHEN p >= 4 THEN 1 ELSE 0 END)  AS promotores,
            SUM(CASE WHEN p  = 3 THEN 1 ELSE 0 END)  AS pasivos,
            SUM(CASE WHEN p <= 2 THEN 1 ELSE 0 END)  AS detractores,
            MAX(Fecha)                               AS ultima,
            MIN(Fecha)                               AS primera
        FROM prom
    """

    # ────────────────────────────────────────────────────────────
    # KPIs OFICINA DIGITAL del agente (escala 1-3)
    # ────────────────────────────────────────────────────────────
    QUERY_KPI_OD = f"""
        WITH cal AS (
            SELECT respuesta_id,
                   {VALOR_OD} AS v
            FROM {VISTA_OD}
            WHERE Agente = %s
            AND Respuesta IN ('Excelente', 'Regular', 'Malo')
        ),
        prom AS (
            SELECT respuesta_id,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id
        )
        SELECT
            COUNT(DISTINCT respuesta_id)              AS total_enc,
            ROUND(AVG(p), 2)                          AS promedio,
            SUM(CASE WHEN p >= 2.5 THEN 1 ELSE 0 END) AS promotores,
            SUM(CASE WHEN p  = 2   THEN 1 ELSE 0 END) AS pasivos,
            SUM(CASE WHEN p <  2   THEN 1 ELSE 0 END) AS detractores
        FROM prom
    """

    # ────────────────────────────────────────────────────────────
    # TENDENCIA MENSUAL — satisfacción + volumen de encuestas
    # ────────────────────────────────────────────────────────────
    QUERY_TENDENCIA = f"""
        WITH cal AS (
            SELECT respuesta_id,
                   YEAR(Fecha)  AS anio,
                   MONTH(Fecha) AS mes,
                   {VALOR_SAT}  AS v
            FROM {VISTA_SAT}
            WHERE Agente = %s
            AND Pregunta IN {PREGUNTAS_SAT}
        ),
        prom AS (
            SELECT respuesta_id, anio, mes,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, anio, mes
        )
        SELECT
            anio, mes,
            COUNT(DISTINCT respuesta_id)             AS total_enc,
            ROUND(AVG(p), 2)                         AS promedio,
            SUM(CASE WHEN p >= 4 THEN 1 ELSE 0 END)  AS promotores,
            SUM(CASE WHEN p  = 3 THEN 1 ELSE 0 END)  AS pasivos,
            SUM(CASE WHEN p <= 2 THEN 1 ELSE 0 END)  AS detractores
        FROM prom
        GROUP BY anio, mes
        ORDER BY anio ASC, mes ASC
    """

    # ────────────────────────────────────────────────────────────
    # RANKING vs pares de la misma unidad
    # ────────────────────────────────────────────────────────────
    QUERY_RANKING = f"""
        WITH cal AS (
            SELECT Agente, respuesta_id,
                   {VALOR_SAT} AS v
            FROM {VISTA_SAT}
            WHERE Unidad = %s
            AND Pregunta IN {PREGUNTAS_SAT}
        ),
        prom AS (
            SELECT Agente, respuesta_id,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY Agente, respuesta_id
        ),
        ranking AS (
            SELECT
                Agente,
                COUNT(DISTINCT respuesta_id)             AS total_enc,
                ROUND(AVG(p), 2)                         AS promedio,
                SUM(CASE WHEN p >= 4 THEN 1 ELSE 0 END)  AS promotores,
                SUM(CASE WHEN p <= 2 THEN 1 ELSE 0 END)  AS detractores,
                RANK() OVER (ORDER BY AVG(p) DESC)        AS posicion
            FROM prom
            GROUP BY Agente
        )
        SELECT *, COUNT(*) OVER() AS total_agentes
        FROM ranking
        ORDER BY posicion ASC
    """

    # ────────────────────────────────────────────────────────────
    # DISTRIBUCIÓN POR TIPO DE GESTIÓN
    # ────────────────────────────────────────────────────────────
    QUERY_DIST_GESTION = f"""
        WITH cal AS (
            SELECT respuesta_id, Gestion,
                   {VALOR_SAT} AS v
            FROM {VISTA_SAT}
            WHERE Agente = %s
            AND Pregunta IN {PREGUNTAS_SAT}
        ),
        prom AS (
            SELECT respuesta_id, Gestion,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, Gestion
        )
        SELECT TOP 8
            Gestion,
            COUNT(DISTINCT respuesta_id)             AS total,
            ROUND(AVG(p), 2)                         AS promedio,
            SUM(CASE WHEN p >= 4 THEN 1 ELSE 0 END)  AS promotores,
            SUM(CASE WHEN p <= 2 THEN 1 ELSE 0 END)  AS detractores
        FROM prom
        GROUP BY Gestion
        ORDER BY total DESC
    """

    # ────────────────────────────────────────────────────────────
    # ÚLTIMAS 10 ENCUESTAS del agente
    # ────────────────────────────────────────────────────────────
    QUERY_ULTIMAS = f"""
        WITH cal AS (
            SELECT respuesta_id, Fecha, Hora, Nombre, Cedula, Gestion, Sucursal,
                   {VALOR_SAT} AS v
            FROM {VISTA_SAT}
            WHERE Agente = %s
            AND Pregunta IN {PREGUNTAS_SAT}
        ),
        prom AS (
            SELECT respuesta_id, Fecha, Hora, Nombre, Cedula, Gestion, Sucursal,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, Fecha, Hora, Nombre, Cedula, Gestion, Sucursal
        )
        SELECT TOP 10
            respuesta_id, Fecha, Hora, Nombre, Cedula, Gestion, Sucursal,
            ROUND(p, 2) AS promedio,
            CASE
                WHEN p >= 4 THEN 'promotor'
                WHEN p  = 3 THEN 'pasivo'
                WHEN p <= 2 THEN 'detractor'
            END AS clasificacion
        FROM prom
        ORDER BY Fecha DESC
    """

    # ────────────────────────────────────────────────────────────
    # OPCIONES PARA FILTROS DEL ÍNDICE
    # ────────────────────────────────────────────────────────────
    QUERY_OPCIONES_UNIDAD = f"""
        SELECT DISTINCT Unidad
        FROM {VISTA_SAT}
        WHERE Unidad IS NOT NULL
        ORDER BY Unidad
    """

    QUERY_OPCIONES_SUCURSAL = f"""
        SELECT DISTINCT Sucursal
        FROM {VISTA_SAT}
        WHERE Sucursal IS NOT NULL
        ORDER BY Sucursal
    """

    # ════════════════════════════════════════════════════════════
    # MÉTODOS PÚBLICOS
    # ════════════════════════════════════════════════════════════

    @classmethod
    def obtener_info(cls, agente: str) -> dict:
        """Info básica del agente (unidad, sucursal) — query liviano."""
        rows = ReportesDBService.ejecutar_query(cls.QUERY_INFO, [agente])
        return rows[0] if rows else {}

    @classmethod
    def obtener_lista(cls, unidad=None, sucursal=None,
                      fecha_inicio=None, fecha_fin=None) -> list:
        conds, params = [], []
        if unidad:
            conds.append("AND Unidad = %s")
            params.append(unidad)
        if sucursal:
            conds.append("AND Sucursal = %s")
            params.append(sucursal)
        if fecha_inicio:
            conds.append("AND Fecha >= %s")
            params.append(fecha_inicio)
        if fecha_fin:
            conds.append("AND Fecha <= %s")
            params.append(fecha_fin + " 23:59:59")
        sql = cls.QUERY_LISTA.format(filtros=" ".join(conds))
        return ReportesDBService.ejecutar_query(sql, params)

    @classmethod
    def obtener_kpi(cls, agente: str,
                        fecha_inicio=None, fecha_fin=None) -> dict:
        conds, params = [], [agente]
        if fecha_inicio:
            conds.append("AND Fecha >= %s")
            params.append(fecha_inicio)
        if fecha_fin:
            conds.append("AND Fecha <= %s")
            params.append(fecha_fin + " 23:59:59")
        sql  = cls.QUERY_KPI_SAT.format(filtro_fecha=" ".join(conds))
        rows = ReportesDBService.ejecutar_query(sql, params)
        if not rows or not rows[0].get("total_enc"):
            return {}
        r     = rows[0]
        total = r["total_enc"] or 1
        p     = r["promedio"]  or 0
        return {
            "total_enc":       r["total_enc"],
            "promedio":        round(p, 2),
            "promedio_pct":    round((p / 5) * 100),
            "promotores":      r["promotores"],
            "promotores_pct":  round((r["promotores"] / total) * 100),
            "pasivos":         r["pasivos"],
            "pasivos_pct":     round((r["pasivos"]    / total) * 100),
            "detractores":     r["detractores"],
            "detractores_pct": round((r["detractores"]/ total) * 100),
            "ultima":          r.get("ultima"),
            "primera":         r.get("primera"),
        }

    @classmethod
    def obtener_kpi_od(cls, agente: str) -> dict:
        """Oficina Digital — escala 1-3. Retorna {} si no hay datos."""
        try:
            rows = ReportesDBService.ejecutar_query(cls.QUERY_KPI_OD, [agente])
            if not rows or not rows[0].get("total_enc"):
                return {}
            r     = rows[0]
            total = r["total_enc"] or 1
            p     = r["promedio"]  or 0
            return {
                "total_enc":       r["total_enc"],
                "promedio":        round(p, 2),
                "promedio_pct":    round((p / 3) * 100),  # escala /3
                "promotores_pct":  round((r["promotores"] / total) * 100),
                "detractores_pct": round((r["detractores"]/ total) * 100),
            }
        except Exception:
            return {}

    @classmethod
    def obtener_tendencia(cls, agente: str) -> list:
        rows = ReportesDBService.ejecutar_query(cls.QUERY_TENDENCIA, [agente])
        for r in rows:
            r["mes_label"]    = f"{MESES[r['mes']]} {r['anio']}"
            r["promedio_pct"] = round((r["promedio"] / 5) * 100) if r["promedio"] else 0
        return rows

    @classmethod
    def obtener_ranking(cls, unidad: str, agente: str) -> dict:
        if not unidad:
            return {"ranking": [], "posicion": None,
                    "total_agentes": 0, "agente_row": None}
        rows      = ReportesDBService.ejecutar_query(cls.QUERY_RANKING, [unidad])
        agente_row = next((r for r in rows if r["Agente"] == agente), None)
        return {
            "ranking":       rows,
            "agente_row":    agente_row,
            "posicion":      agente_row["posicion"]    if agente_row else None,
            "total_agentes": rows[0]["total_agentes"]  if rows       else 0,
        }

    @classmethod
    def obtener_dist_gestion(cls, agente: str) -> list:
        return ReportesDBService.ejecutar_query(cls.QUERY_DIST_GESTION, [agente])

    @classmethod
    def obtener_ultimas(cls, agente: str) -> list:
        return ReportesDBService.ejecutar_query(cls.QUERY_ULTIMAS, [agente])

    @classmethod
    def obtener_opciones(cls) -> dict:
        return {
            "unidades":   ReportesDBService.ejecutar_query(cls.QUERY_OPCIONES_UNIDAD),
            "sucursales": ReportesDBService.ejecutar_query(cls.QUERY_OPCIONES_SUCURSAL),
        }