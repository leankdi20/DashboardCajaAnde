# ═══════════════════════════════════════════════════════════════════
# apps/dashboard/reports/reporte_perfil_agente_od.py
# Perfil de agentes — Oficina Digital (escala 1-3)
#
# Columnas disponibles en VISTA_OD:
#   respuesta_id, encuesta_id, encuesta_nombre, encuesta_correos,
#   encuesta_creado, Fecha, Cedula, Nombre, Agente,
#   orden, encuesta_det_id, Pregunta, Respuesta
#
# NO existen: Unidad, Sucursal, Gestion
# Se simulan con NULL AS / literales para mantener interfaz uniforme
# ═══════════════════════════════════════════════════════════════════

from ..services.db_service import ReportesDBService

# ── Vista SQL centralizada ────────────────────────────────────────
VISTA_OD = "dbo.vw_reporte_encuestas_satisfaccion_oficina_digital"

# ── Lógica de valoración OD (escala 1-3) ─────────────────────────
VALOR_OD = """
    CASE Respuesta
        WHEN 'Excelente' THEN 3
        WHEN 'Regular'   THEN 2
        WHEN 'Malo'      THEN 1
        ELSE NULL
    END
"""

MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


class ReportePerfilAgenteOD:

    # ────────────────────────────────────────────────────────────
    # INFO RÁPIDA DEL AGENTE
    # No hay Unidad/Sucursal — se retornan como NULL
    # ────────────────────────────────────────────────────────────
    QUERY_INFO = f"""
        SELECT TOP 1
            Agente,
            'Oficina Digital' AS Unidad,
            NULL              AS Sucursal
        FROM {VISTA_OD}
        WHERE Agente = %s
    """

    # ────────────────────────────────────────────────────────────
    # LISTA DE AGENTES con métricas OD
    # Ranking global (sin PARTITION BY Unidad)
    # ────────────────────────────────────────────────────────────
    QUERY_LISTA = f"""
        WITH cal AS (
            SELECT Agente, respuesta_id,
                {VALOR_OD} AS v
            FROM {VISTA_OD}
            WHERE Respuesta IN ('Excelente', 'Regular', 'Malo')
            {{filtros}}
        ),
        prom AS (
            SELECT Agente, respuesta_id,
                AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY Agente, respuesta_id
        )
        SELECT
            Agente,
            'Oficina Digital'                               AS Unidad,
            NULL                                            AS Sucursal,
            COUNT(DISTINCT respuesta_id)                    AS total_enc,
            ROUND(AVG(p), 2)                                AS promedio,
            SUM(CASE WHEN p >= 2.5 THEN 1 ELSE 0 END)       AS promotores,
            SUM(CASE WHEN p  = 2   THEN 1 ELSE 0 END)       AS pasivos,
            SUM(CASE WHEN p <  2   THEN 1 ELSE 0 END)       AS detractores,
            RANK() OVER (ORDER BY AVG(p) DESC)               AS rank_unidad
        FROM prom
        WHERE Agente IS NOT NULL
        GROUP BY Agente
        ORDER BY AVG(p) DESC
    """

    # ────────────────────────────────────────────────────────────
    # KPIs del agente en OD
    # ────────────────────────────────────────────────────────────
    QUERY_KPI = f"""
        WITH cal AS (
            SELECT respuesta_id, Fecha,
                   {VALOR_OD} AS v
            FROM {VISTA_OD}
            WHERE Agente = %s
            AND Respuesta IN ('Excelente', 'Regular', 'Malo')
            {{filtro_fecha}}
        ),
        prom AS (
            SELECT respuesta_id, Fecha,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, Fecha
        )
        SELECT
            COUNT(DISTINCT respuesta_id)              AS total_enc,
            ROUND(AVG(p), 2)                          AS promedio,
            SUM(CASE WHEN p >= 2.5 THEN 1 ELSE 0 END) AS promotores,
            SUM(CASE WHEN p  = 2   THEN 1 ELSE 0 END) AS pasivos,
            SUM(CASE WHEN p <  2   THEN 1 ELSE 0 END) AS detractores,
            MAX(Fecha)                                AS ultima,
            MIN(Fecha)                                AS primera
        FROM prom
    """

    # ────────────────────────────────────────────────────────────
    # TENDENCIA MENSUAL
    # ────────────────────────────────────────────────────────────
    QUERY_TENDENCIA = f"""
        WITH cal AS (
            SELECT respuesta_id,
                   YEAR(Fecha)  AS anio,
                   MONTH(Fecha) AS mes,
                   {VALOR_OD}   AS v
            FROM {VISTA_OD}
            WHERE Agente = %s
            AND Respuesta IN ('Excelente', 'Regular', 'Malo')
        ),
        prom AS (
            SELECT respuesta_id, anio, mes,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, anio, mes
        )
        SELECT
            anio, mes,
            COUNT(DISTINCT respuesta_id)              AS total_enc,
            ROUND(AVG(p), 2)                          AS promedio,
            SUM(CASE WHEN p >= 2.5 THEN 1 ELSE 0 END) AS promotores,
            SUM(CASE WHEN p  = 2   THEN 1 ELSE 0 END) AS pasivos,
            SUM(CASE WHEN p <  2   THEN 1 ELSE 0 END) AS detractores
        FROM prom
        GROUP BY anio, mes
        ORDER BY anio ASC, mes ASC
    """

    # ────────────────────────────────────────────────────────────
    # RANKING GLOBAL — sin Unidad, rankea todos los agentes de OD
    # El parámetro unidad se ignora (no existe en la vista)
    # ────────────────────────────────────────────────────────────
    QUERY_RANKING = f"""
        WITH cal AS (
            SELECT Agente, respuesta_id,
                   {VALOR_OD} AS v
            FROM {VISTA_OD}
            WHERE Respuesta IN ('Excelente', 'Regular', 'Malo')
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
                COUNT(DISTINCT respuesta_id)              AS total_enc,
                ROUND(AVG(p), 2)                          AS promedio,
                SUM(CASE WHEN p >= 2.5 THEN 1 ELSE 0 END) AS promotores,
                SUM(CASE WHEN p <  2   THEN 1 ELSE 0 END) AS detractores,
                RANK() OVER (ORDER BY AVG(p) DESC)         AS posicion
            FROM prom
            GROUP BY Agente
        )
        SELECT *, COUNT(*) OVER() AS total_agentes
        FROM ranking
        ORDER BY posicion ASC
    """

    # ────────────────────────────────────────────────────────────
    # DISTRIBUCIÓN POR PREGUNTA (equivalente a "gestión" en SAT)
    # No hay Gestion — se usa Pregunta como agrupación
    # ────────────────────────────────────────────────────────────
    QUERY_DIST_GESTION = f"""
        WITH cal AS (
            SELECT respuesta_id, Pregunta,
                   {VALOR_OD} AS v
            FROM {VISTA_OD}
            WHERE Agente = %s
            AND Respuesta IN ('Excelente', 'Regular', 'Malo')
        ),
        prom AS (
            SELECT respuesta_id, Pregunta,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, Pregunta
        )
        SELECT TOP 8
            Pregunta                                   AS Gestion,
            COUNT(DISTINCT respuesta_id)               AS total,
            ROUND(AVG(p), 2)                           AS promedio,
            SUM(CASE WHEN p >= 2.5 THEN 1 ELSE 0 END)  AS promotores,
            SUM(CASE WHEN p <  2   THEN 1 ELSE 0 END)  AS detractores
        FROM prom
        GROUP BY Pregunta
        ORDER BY total DESC
    """

    # ────────────────────────────────────────────────────────────
    # ÚLTIMAS 10 ENCUESTAS
    # Gestion y Sucursal no existen — se retornan como NULL
    # ────────────────────────────────────────────────────────────
    QUERY_ULTIMAS = f"""
        WITH cal AS (
            SELECT respuesta_id, Fecha, Nombre, Cedula,
                   {VALOR_OD} AS v
            FROM {VISTA_OD}
            WHERE Agente = %s
            AND Respuesta IN ('Excelente', 'Regular', 'Malo')
        ),
        prom AS (
            SELECT respuesta_id, Fecha, Nombre, Cedula,
                   AVG(CAST(v AS FLOAT)) AS p
            FROM cal WHERE v IS NOT NULL
            GROUP BY respuesta_id, Fecha, Nombre, Cedula
        )
        SELECT TOP 10
            respuesta_id,
            Fecha,
            Nombre,
            Cedula,
            NULL        AS Gestion,
            NULL        AS Sucursal,
            ROUND(p, 2) AS promedio,
            CASE
                WHEN p >= 2.5 THEN 'promotor'
                WHEN p  = 2   THEN 'pasivo'
                WHEN p <  2   THEN 'detractor'
            END AS clasificacion
        FROM prom
        ORDER BY Fecha DESC
    """

    # ────────────────────────────────────────────────────────────
    # OPCIONES PARA FILTROS DEL ÍNDICE
    # No hay Unidad/Sucursal — se retorna literal/vacío
    # ────────────────────────────────────────────────────────────
    QUERY_OPCIONES_UNIDAD = """
        SELECT 'Oficina Digital' AS Unidad
    """

    QUERY_OPCIONES_SUCURSAL = f"""
        SELECT DISTINCT NULL AS Sucursal
        FROM {VISTA_OD}
        WHERE 1 = 0
    """

    # ════════════════════════════════════════════════════════════
    # MÉTODOS PÚBLICOS — misma interfaz que ReportePerfilAgente
    # ════════════════════════════════════════════════════════════

    @classmethod
    def obtener_info(cls, agente: str) -> dict:
        rows = ReportesDBService.ejecutar_query(cls.QUERY_INFO, [agente])
        return rows[0] if rows else {}

    @classmethod
    def obtener_lista(cls, unidad=None, sucursal=None,
                      fecha_inicio=None, fecha_fin=None) -> list:
        # OD no tiene Unidad/Sucursal — solo fecha aplica como filtro
        conds, params = [], []
        if fecha_inicio:
            conds.append("AND Fecha >= %s"); params.append(fecha_inicio)
        if fecha_fin:
            conds.append("AND Fecha <= %s"); params.append(fecha_fin + " 23:59:59")
        sql = cls.QUERY_LISTA.format(filtros=" ".join(conds))
        return ReportesDBService.ejecutar_query(sql, params)

    @classmethod
    def obtener_kpi(cls, agente: str,
                    fecha_inicio=None, fecha_fin=None) -> dict:
        conds, params = [], [agente]
        if fecha_inicio:
            conds.append("AND Fecha >= %s"); params.append(fecha_inicio)
        if fecha_fin:
            conds.append("AND Fecha <= %s"); params.append(fecha_fin + " 23:59:59")
        sql  = cls.QUERY_KPI.format(filtro_fecha=" ".join(conds))
        rows = ReportesDBService.ejecutar_query(sql, params)
        if not rows or not rows[0].get("total_enc"):
            return {}
        r     = rows[0]
        total = r["total_enc"] or 1
        p     = r["promedio"]  or 0
        return {
            "total_enc":       r["total_enc"],
            "promedio":        round(p, 2),
            "promedio_pct":    round((p / 3) * 100),   # escala /3
            "promotores":      r["promotores"],
            "promotores_pct":  round((r["promotores"]  / total) * 100),
            "pasivos":         r["pasivos"],
            "pasivos_pct":     round((r["pasivos"]      / total) * 100),
            "detractores":     r["detractores"],
            "detractores_pct": round((r["detractores"]  / total) * 100),
            "ultima":          r.get("ultima"),
            "primera":         r.get("primera"),
            "escala":          "1-3 (Excelente / Regular / Malo)",
        }

    @classmethod
    def obtener_tendencia(cls, agente: str) -> list:
        rows = ReportesDBService.ejecutar_query(cls.QUERY_TENDENCIA, [agente])
        for r in rows:
            r["mes_label"]    = f"{MESES[r['mes']]} {r['anio']}"
            r["promedio_pct"] = round((r["promedio"] / 3) * 100) if r["promedio"] else 0
        return rows

    @classmethod
    def obtener_ranking(cls, unidad: str, agente: str) -> dict:
        # OD no tiene unidades — rankea globalmente todos los agentes
        rows      = ReportesDBService.ejecutar_query(cls.QUERY_RANKING)
        agente_row = next((r for r in rows if r["Agente"] == agente), None)
        return {
            "ranking":       rows,
            "agente_row":    agente_row,
            "posicion":      agente_row["posicion"]   if agente_row else None,
            "total_agentes": rows[0]["total_agentes"] if rows       else 0,
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