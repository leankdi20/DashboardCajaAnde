from ..services.db_service import ReportesDBService

class ReporteEncuestaFeriaSalud:

    VISTA = "dbo.vw_reporte_encuestas_satisfaccion_Feria_de_la_salud"

    QUERY = f"""
        SELECT respuesta_id, Fecha, Pregunta, Respuesta
        FROM {VISTA}
        ORDER BY Fecha DESC
    """

    QUERY_FILTRADO = f"""
        SELECT respuesta_id, Fecha, Pregunta, Respuesta
        FROM {VISTA}
        WHERE 1=1
        {{filtros_base}}
        ORDER BY Fecha DESC
    """

    QUERY_TOTAL = f"""
        SELECT COUNT(DISTINCT respuesta_id) as total_encuestas
        FROM {VISTA}
        {{filtros}}
    """

    @classmethod
    def obtener_datos(cls, filtros: dict = None) -> list[dict]:
        if not filtros or not any(filtros.values()):
            return ReportesDBService.ejecutar_query(cls.QUERY)

        condiciones_base = []
        params = []

        if filtros.get("fecha_inicio"):
            condiciones_base.append("AND CAST(Fecha AS DATE) >= %s")
            params.append(filtros["fecha_inicio"])
        if filtros.get("fecha_fin"):
            condiciones_base.append("AND CAST(Fecha AS DATE) <= %s")
            params.append(filtros["fecha_fin"])
        if filtros.get("pregunta"):
            condiciones_base.append("AND Pregunta = %s")
            params.append(filtros["pregunta"])
        if filtros.get("respuesta"):
            condiciones_base.append("AND Respuesta = %s")
            params.append(filtros["respuesta"])

        sql = cls.QUERY_FILTRADO.format(filtros_base=" ".join(condiciones_base))
        return ReportesDBService.ejecutar_query(sql, params)
    
    @classmethod
    def obtener_total(cls, fecha_inicio: str = None, fecha_fin: str = None) -> int:
        condiciones = []
        params = []

        if fecha_inicio:
            condiciones.append("AND CAST(Fecha AS DATE) >= %s")
            params.append(fecha_inicio)
        if fecha_fin:
            condiciones.append("AND CAST(Fecha AS DATE) <= %s")
            params.append(fecha_fin)

        filtros_sql = "WHERE 1=1 " + " ".join(condiciones) if condiciones else ""
        sql = cls.QUERY_TOTAL.format(filtros=filtros_sql)
        resultado = ReportesDBService.ejecutar_query(sql, params)
        return resultado[0].get("total_encuestas", 0) if resultado else 0
    

    @classmethod
    def obtener_datos_agrupados(cls, filtros: dict = None) -> tuple[list[dict], list[str]]:
        if not filtros or not any(filtros.values()):
            raw = ReportesDBService.ejecutar_query(cls.QUERY)
        else:
            condiciones_base = []
            params = []
            if filtros.get("fecha_inicio"):
                condiciones_base.append("AND CAST(Fecha AS DATE) >= %s")
                params.append(filtros["fecha_inicio"])
            if filtros.get("fecha_fin"):
                condiciones_base.append("AND CAST(Fecha AS DATE) <= %s")
                params.append(filtros["fecha_fin"])
            sql = cls.QUERY_FILTRADO.format(filtros_base=" ".join(condiciones_base))
            raw = ReportesDBService.ejecutar_query(sql, params)

        encuestas = {}
        preguntas_orden = []

        for fila in raw:
            rid       = fila["respuesta_id"]
            pregunta  = fila.get("Pregunta", "")
            respuesta = fila.get("Respuesta", "")

            if rid not in encuestas:
                encuestas[rid] = {
                    "respuesta_id": rid,
                    "Fecha":        fila.get("Fecha", ""),
                }
            if pregunta and pregunta not in preguntas_orden:
                preguntas_orden.append(pregunta)
            encuestas[rid][pregunta] = respuesta

        return list(encuestas.values()), preguntas_orden