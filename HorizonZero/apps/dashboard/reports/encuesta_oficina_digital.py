from ..services.db_service import ReportesDBService


class ReporteEncuestaOficinaDigital:

    VISTA = "dbo.vw_reporte_encuestas_satisfaccion_oficina_digital"

    QUERY = f"SELECT * FROM {VISTA} ORDER BY Fecha DESC"

    QUERY_FILTRADO = f"""
        SELECT * FROM {VISTA}
        WHERE 1=1
        {{filtros}}
        ORDER BY Fecha DESC
    """

    QUERY_DETALLE = f"""
        SELECT * FROM {VISTA}
        WHERE respuesta_id = %s
        ORDER BY orden
    """

    QUERY_PROMEDIO_AGENTE = f"""
        SELECT 
            COUNT(DISTINCT respuesta_id) as total_encuestas,
            AVG(CAST(
                CASE 
                    WHEN Pregunta = '¿Cómo valora la atención que le brindó el ejecutivo de servicio?' THEN
                        CASE Respuesta
                            WHEN 'Excelente' THEN 3
                            WHEN 'Regular'   THEN 2
                            WHEN 'Malo'      THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo valora el espacio físico asignado para la oficina digital?' THEN
                        CASE Respuesta
                            WHEN 'Excelente' THEN 3
                            WHEN 'Regular'   THEN 2
                            WHEN 'Malo'      THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo califica la velocidad de internet para realizar sus trámites en la Oficina Digital?' THEN
                        CASE Respuesta
                            WHEN 'Excelente' THEN 3
                            WHEN 'Regular'   THEN 2
                            WHEN 'Malo'      THEN 1
                            ELSE NULL
                        END
                    WHEN Pregunta = '¿Cómo fue su experiencia al realizar sus gestiones en la Oficina Digital?' THEN
                        CASE Respuesta
                            WHEN 'Excelente' THEN 3
                            WHEN 'Regular'   THEN 2
                            WHEN 'Malo'      THEN 1
                            ELSE NULL
                        END
                    ELSE NULL
                END AS FLOAT
            )) as promedio_agente
        FROM {VISTA}
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
            )) as promedio_encuesta
        FROM {VISTA}
        WHERE respuesta_id = %s
    """

    QUERY_KPIS_GLOBALES = f"""
        SELECT
            COUNT(DISTINCT respuesta_id) as total_encuestas,
            AVG(CAST(
                CASE Respuesta
                        WHEN 'Excelente' THEN 3
                        WHEN 'Regular'   THEN 2
                        WHEN 'Malo'      THEN 1
                    ELSE NULL
                END AS FLOAT
            )) as promedio_general
        FROM {VISTA}
        WHERE Respuesta IN ('Excelente','Regular','Malo')
        {{filtros}}
    """

    @classmethod
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        if not filtros or not any(filtros.values()):
            return ReportesDBService.ejecutar_query(cls.QUERY)

        condiciones = []
        params = []

        if filtros.get("agente"):
            condiciones.append("AND Agente = %s")
            params.append(filtros["agente"])
        if filtros.get("nombre"):
            condiciones.append("AND Nombre = %s")
            params.append(filtros["nombre"])
        if filtros.get("cedula"):
            condiciones.append("AND Cedula = %s")
            params.append(filtros["cedula"])
        if filtros.get("fecha_inicio"):
            condiciones.append("AND Fecha >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones.append("AND Fecha <= %s")
            params.append(filtros["fecha_fin"] + " 23:59:59")

        sql = cls.QUERY_FILTRADO.format(filtros=" ".join(condiciones))
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
    def obtener_kpis_globales(cls, sucursales: list = None, fecha_inicio: str = None, fecha_fin: str = None) -> dict:
        condiciones = []
        params = []

        
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
        promedio = r.get("promedio_general") or 0
        return {
            "total_encuestas":  r.get("total_encuestas", 0),
            "promedio_general": round((promedio / 3) * 100) if promedio else 0,  # ← /3 no /5
        }