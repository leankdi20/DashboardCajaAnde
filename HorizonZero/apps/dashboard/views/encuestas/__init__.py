# apps/dashboard/views/encuestas/__init__.py
from .satisfaccion import (
    encuesta_satisfaccion,
    encuesta_satisfaccion_buscar,
    encuesta_satisfaccion_kpis,
    encuesta_satisfaccion_detalle,
    encuesta_satisfaccion_exportar,
    encuesta_satisfaccion_detalle_exportar,
)
from .oficina_digital import (
    encuesta_satisfaccion_oficina,
    encuesta_satisfaccion_detalle_of_dig,
    encuesta_satisfaccion_of_dig_exportar,
    encuesta_satisfaccion_detalle_of_dig_exportar,
    encuesta_oficina_digital_buscar,
    encuesta_oficina_digital_kpis,
)
from .whatsapp_agente import (
    encuesta_whatsApp_agente,
    encuesta_whatsapp_agente_detalle,
    encuesta_whatsapp_agente_exportar,
    encuesta_whatsapp_agente_detalle_exportar,
)
from .whatsapp import (
    encuesta_whatsApp_,
    encuesta_whatsapp_detalle,
    encuesta_whatsapp_exportar,
    encuesta_whatsapp_detalle_exportar,
)
from .pagina_web import (
    encuesta_experiencia_web,
    encuesta_experiencia_web_detalle,
    encuesta_experiencia_web_exportar,
)
from .feria_salud import (
    encuesta_feria_salud_,
    encuesta_feria_salud_exportar,
)
from .perfil_agentes import (
    perfil_agentes_sat_index,
    perfil_agentes_sat_data,
    perfil_agente_sat_detalle,
    perfil_agente_sat_ajax,
    perfil_agentes_od_index,
    perfil_agentes_od_data,
    perfil_agente_od_detalle,
    perfil_agente_od_ajax,
)
