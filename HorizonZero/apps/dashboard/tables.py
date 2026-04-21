from django.urls import reverse
from django.utils.safestring import mark_safe
import django_tables2 as tables

# ENCUESTA SATISFACCION
class EncuestaSatisfaccionTable(tables.Table):
    Fecha      = tables.Column(verbose_name="Fecha")
    Agente     = tables.Column(verbose_name="Agente")
    Unidad     = tables.Column(verbose_name="Unidad")
    Sucursal   = tables.Column(verbose_name="Sucursal")
    Gestion    = tables.Column(verbose_name="Gestión")
    Nombre     = tables.Column(verbose_name="Accionista")
    Cedula     = tables.Column(verbose_name="Cédula")
    acciones   = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")

    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/encuestas/satisfaccion/{record["respuesta_id"]}/" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )

    class Meta:
        template_name = "dashboard/components/table.html"
        attrs = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence = ("Fecha", "Agente", "Unidad", "Sucursal", "Gestion", "Nombre", "Cedula", "acciones")



# OFICINA DIGITAL

class EncuestaOficinaDigitalTable(tables.Table):
    Fecha    = tables.Column(verbose_name="Fecha")
    Cedula   = tables.Column(verbose_name="Cédula")
    Nombre   = tables.Column(verbose_name="Accionista")
    Agente   = tables.Column(verbose_name="Agente")
    acciones = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")

    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/encuestas/satisfaccion_of_dig/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )

    class Meta:
        template_name = "dashboard/components/table.html"
        attrs         = {"class": "w-full text-left border-collapse"}
        row_attrs     = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence      = ("Fecha", "Cedula", "Nombre", "Agente", "acciones")




# EXPERIENCIA PÁGINA WEB
class EncuestaPaginaWebTable(tables.Table):
    Fecha = tables.Column(verbose_name="Fecha")
    Nombre = tables.Column(verbose_name="Nombre")
    acciones = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")

    def render_acciones(self, record):
        respuesta_id = record.get("respuesta_id")

        if not respuesta_id:
            return mark_safe(
                '<span class="text-slate-400" title="Sin respuesta_id">'
                '<span class="material-symbols-outlined">visibility_off</span>'
                '</span>'
            )

        url = reverse("dashboard:encuesta_experiencia_web_detalle", args=[respuesta_id])

        return mark_safe(
            f'<a href="{url}" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors" '
            f'title="Ver detalle">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )

    class Meta:
        template_name = "dashboard/components/table.html"
        attrs = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence = ("Fecha", "Nombre", "acciones")



# Encuesta WhatsApp Agente
class EncuestaWhatsappAgenteTable(tables.Table):
    Fecha    = tables.Column(verbose_name="Fecha")
    Nombre   = tables.Column(verbose_name="Accionista")
    Cedula   = tables.Column(verbose_name="Cédula")
    Agente   = tables.Column(verbose_name="Agente")
    acciones = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")

    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/encuestas/whatsapp-agente/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )

    class Meta:
        template_name = "dashboard/components/table.html"
        attrs         = {"class": "w-full text-left border-collapse"}
        row_attrs     = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence      = ("Fecha", "Agente", "Nombre", "Cedula", "acciones")



# Encuesta WhatsApp
class EncuestaWhatsappTable(tables.Table):
    Fecha    = tables.Column(verbose_name="Fecha")
    Nombre   = tables.Column(verbose_name="Accionista")
    Cedula   = tables.Column(verbose_name="Cédula")
    acciones = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")

    def render_acciones(self, record):
        url = reverse("dashboard:encuesta_whatsapp_detalle", args=[record["respuesta_id"]])
        return mark_safe(
            f'<a href="{url}" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )

    class Meta:
        template_name = "dashboard/components/table.html"
        attrs         = {"class": "w-full text-left border-collapse"}
        row_attrs     = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence      = ("Fecha", "Nombre", "Cedula", "acciones")




#Encuesta Feria de la Salud
class EncuestaFeriaSaludTable(tables.Table):
    Fecha     = tables.Column(verbose_name="Fecha")
    Pregunta  = tables.Column(verbose_name="Pregunta")
    Respuesta = tables.Column(verbose_name="Respuesta")
    

    def render_acciones(self, record):
        url = reverse("dashboard:encuesta_feria_salud_detalle", args=[record["respuesta_id"]])
        return mark_safe(
            f'<a href="{url}" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )
    


    class Meta:
        template_name = "dashboard/components/table.html"
        attrs         = {"class": "w-full text-left border-collapse"}
        row_attrs     = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence      = ("Fecha", "Pregunta", "Respuesta")



class SolicitudTarjetaCreditoTable(tables.Table):
    FechaHora   = tables.Column(verbose_name="Fecha / Hora")
    Cedula      = tables.Column(verbose_name="Cédula")
    Nombre      = tables.Column(verbose_name="Nombre")
    Telefono    = tables.Column(verbose_name="Teléfono")
    Correo      = tables.Column(verbose_name="Correo")
    TipoTramite = tables.Column(verbose_name="Tipo de Trámite")
    acciones    = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/tarjetas/solicitudes/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "Nombre", "Telefono", "Correo", "TipoTramite", "acciones")
 




class SolicitudTarjetaDebito_CiuOroTable(tables.Table):
    FechaHora    = tables.Column(verbose_name="Fecha / Hora")
    Cedula       = tables.Column(verbose_name="Cédula")
    Nombre       = tables.Column(verbose_name="Nombre")
    Correo       = tables.Column(verbose_name="Correo")
    Telefono     = tables.Column(verbose_name="Teléfono")
    Correo       = tables.Column(verbose_name="Correo")
    DireccionEnvio = tables.Column(verbose_name="Dirección de Envío")
    acciones     = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/tarjetas/debito/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "Nombre","Correo", "Telefono",
                     "DireccionEnvio", "acciones")




 
class SolicitudTarjetaDebitoGestionTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    Nombre         = tables.Column(verbose_name="Nombre")
    Telefono       = tables.Column(verbose_name="Teléfono")
    Correo         = tables.Column(verbose_name="Correo")
    TipoTramite    = tables.Column(verbose_name="Tipo de Trámite")
    DireccionEnvio = tables.Column(verbose_name="Dirección de Envío")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/tarjetas/debito-gestion/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "Nombre", "Telefono",
                     "Correo", "TipoTramite", "DireccionEnvio", "acciones")
 


class SolicitudRedencionPuntosTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    Nombre         = tables.Column(verbose_name="Nombre")
    Telefono       = tables.Column(verbose_name="Teléfono")
    Correo         = tables.Column(verbose_name="Correo")
    TipoRedencion  = tables.Column(verbose_name="Tipo de Redención")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/tarjetas/redencion-puntos/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">visibility</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "Nombre", "Telefono",
                     "Correo", "TipoRedencion", "acciones")
 



 
class CajaAndeAsistenciaTable(tables.Table):
    FechaHora    = tables.Column(verbose_name="Fecha / Hora")
    Cedula       = tables.Column(verbose_name="Cédula")
    Nombre       = tables.Column(verbose_name="Nombre")
    Correo       = tables.Column(verbose_name="Correo")
    TipoPlan     = tables.Column(verbose_name="Plan")
    TipoTarjeta  = tables.Column(verbose_name="Tipo Tarjeta")
    acciones     = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/tarjetas/asistencia/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "Nombre", "Correo",
                     "TipoPlan", "TipoTarjeta", "acciones")
 


 
class SolicitudDepositoSalarioTable(tables.Table):
    FechaHora = tables.Column(verbose_name="Fecha / Hora")
    Cedula    = tables.Column(verbose_name="Cédula")
    acciones  = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/ahorros/deposito-salario/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "acciones")


class SolicitudAhorroModCuotaTable(tables.Table):
    FechaHora        = tables.Column(verbose_name="Fecha / Hora")
    Cedula           = tables.Column(verbose_name="Cédula")
    Nombre_Completo  = tables.Column(verbose_name="Nombre")
    TipoAhorro       = tables.Column(verbose_name="Tipo de Ahorro")
    NumeroContrato   = tables.Column(verbose_name="N° Contrato")
    TipoModificacion = tables.Column(verbose_name="Modificación")
    MontoCuotaDeducir = tables.Column(verbose_name="Monto (₡)")
    acciones         = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_MontoCuotaDeducir(self, value):
        if value is None:
            return "—"
        try:
            return f"₡{int(value):,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(value)
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/ahorros/modificacion-cuota/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "Nombre_Completo", "TipoAhorro",
                     "NumeroContrato", "TipoModificacion", "MontoCuotaDeducir", "acciones")
        

class SolicitudReinversionAhorroTable(tables.Table):
    Fecha              = tables.Column(verbose_name="Fecha")
    Hora               = tables.Column(verbose_name="Hora")
    Cedula             = tables.Column(verbose_name="Cédula")
    Nombre_Completo    = tables.Column(verbose_name="Nombre")
    SolicitoReinversion = tables.Column(verbose_name="Tipo de Ahorro")
    TipoReinversion    = tables.Column(verbose_name="Reinversión")
    NumeroContrato     = tables.Column(verbose_name="N° Contrato")
    acciones           = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/ahorros/reinversion/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("Fecha", "Hora", "Cedula", "Nombre_Completo",
                     "SolicitoReinversion", "TipoReinversion", "NumeroContrato", "acciones")
        


class SolicitudAutorizacionAhorroNuevoTable(tables.Table):
    Fecha           = tables.Column(verbose_name="Fecha")
    Cedula          = tables.Column(verbose_name="Cédula")
    Nombre_Completo = tables.Column(verbose_name="Nombre")
    TipoAhorro      = tables.Column(verbose_name="Tipo de Ahorro")
    FormaPago       = tables.Column(verbose_name="Forma de Pago")
    TipoReinversion = tables.Column(verbose_name="Reinversión")
    MontoCuota      = tables.Column(verbose_name="Monto Cuota (₡)")
    acciones        = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_MontoCuota(self, value):
        if value is None:
            return "—"
        try:
            return f"₡{int(value):,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(value)
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/ahorros/autorizacion-nuevo/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("Fecha", "Cedula", "Nombre_Completo", "TipoAhorro",
                     "FormaPago", "TipoReinversion", "MontoCuota", "acciones")
 




class SolicitudCompraVehiculoTable(tables.Table):
    FechaHora             = tables.Column(verbose_name="Fecha / Hora")
    Cedula                = tables.Column(verbose_name="Cédula")
    NombreCompleto        = tables.Column(verbose_name="Nombre")
    TelefonoCelular       = tables.Column(verbose_name="Teléfono")
    TipoVehiculo          = tables.Column(verbose_name="Tipo Vehículo")
    MontoCreditoSolicitado = tables.Column(verbose_name="Monto (₡)")
    Garantia              = tables.Column(verbose_name="Garantía")
    acciones              = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_MontoCreditoSolicitado(self, value):
        if value is None:
            return "—"
        try:
            return f"₡{int(float(value)):,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(value)
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/vivienda/vehiculo/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto", "TelefonoCelular",
                     "TipoVehiculo", "MontoCreditoSolicitado", "Garantia", "acciones")
        



class SolicitudPrestamoViviendaTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    NombreCompleto = tables.Column(verbose_name="Nombre")
    Telefono       = tables.Column(verbose_name="Teléfono")
    TipoPrestamo   = tables.Column(verbose_name="Tipo de Préstamo")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/vivienda/prestamo/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto", "Telefono", "TipoPrestamo", "acciones")
 


class SolicitudPrestamoDesarrolloTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    NombreCompleto = tables.Column(verbose_name="Nombre")
    Telefono       = tables.Column(verbose_name="Teléfono")
    PlanInversion  = tables.Column(verbose_name="Plan de Inversión")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/vivienda/desarrollo/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto", "Telefono", "PlanInversion", "acciones")





class SolicitudPresolicitudCreditoPersonalTable(tables.Table):
    FechaHora             = tables.Column(verbose_name="Fecha / Hora")
    Cedula                = tables.Column(verbose_name="Cédula")
    NombreCompleto        = tables.Column(verbose_name="Nombre")
    Telefono              = tables.Column(verbose_name="Teléfono")
    TipoCredito           = tables.Column(verbose_name="Tipo de Crédito")
    SucursalFormalizacion = tables.Column(verbose_name="Sucursal")
    Monto                 = tables.Column(verbose_name="Monto (₡)")
    acciones              = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_Monto(self, value):
        if value is None:
            return "—"
        try:
            return f"₡{int(float(value)):,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(value)
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/prestamos/credito-personal/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto", "Telefono",
                     "TipoCredito", "SucursalFormalizacion", "Monto", "acciones")
        


class ComprobanteAutorizacionAhorroTable(tables.Table):
    Fecha             = tables.Column(verbose_name="Fecha")
    Hora              = tables.Column(verbose_name="Hora")
    Cedula            = tables.Column(verbose_name="Cédula")
    NombreCompleto    = tables.Column(verbose_name="Nombre")
    NumeroTelefonico  = tables.Column(verbose_name="Teléfono")
    DetallePago       = tables.Column(verbose_name="Detalle de Pago")
    acciones          = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/control-credito/autorizacion-ahorro/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("Fecha", "Hora", "Cedula", "NombreCompleto",
                     "NumeroTelefonico", "DetallePago", "acciones")
 
 
class ComprobantesPagoTable(tables.Table):
    Fecha            = tables.Column(verbose_name="Fecha")
    Hora             = tables.Column(verbose_name="Hora")
    Cedula           = tables.Column(verbose_name="Cédula")
    NombreCompleto   = tables.Column(verbose_name="Nombre")
    Banco            = tables.Column(verbose_name="Banco")
    NumeroDeposito   = tables.Column(verbose_name="N° Depósito")
    Monto            = tables.Column(verbose_name="Monto (₡)")
    acciones         = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_Monto(self, value):
        if value is None:
            return "—"
        try:
            return f"₡{int(float(value)):,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(value)
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/control-credito/comprobantes-pago/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("Fecha", "Hora", "Cedula", "NombreCompleto",
                     "Banco", "NumeroDeposito", "Monto", "acciones")
        


class SolicitudClaveTemporalCajaTelTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    NombreCompleto = tables.Column(verbose_name="Nombre")
    CorreoPersonal = tables.Column(verbose_name="Correo")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/servicio-accionista/clave-cajatel/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto", "CorreoPersonal", "acciones")



class SolicitudSeguroViajeroTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    NombreCompleto = tables.Column(verbose_name="Nombre")
    Correo         = tables.Column(verbose_name="Correo")
    Telefono       = tables.Column(verbose_name="Teléfono")
    Destino        = tables.Column(verbose_name="Destino")
    FechaInicioViaje = tables.Column(verbose_name="Inicio Viaje")
    FechaFinalViaje  = tables.Column(verbose_name="Fin Viaje")
    Parentesco     = tables.Column(verbose_name="Parentesco")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/seguros/viajero/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto", "Correo", "Telefono", "Destino",
                     "FechaInicioViaje", "FechaFinalViaje", "Parentesco", "acciones")
        



class SolicitudMarchamoTable(tables.Table):
    FechaHora      = tables.Column(verbose_name="Fecha / Hora")
    Cedula         = tables.Column(verbose_name="Cédula")
    NombreCompleto = tables.Column(verbose_name="Nombre")
    TipoVehiculo   = tables.Column(verbose_name="Tipo Vehículo")
    NumeroPlaca    = tables.Column(verbose_name="Placa")
    SucursalRetiro = tables.Column(verbose_name="Sucursal")
    acciones       = tables.Column(empty_values=(), orderable=False, verbose_name="Acciones")
 
    def render_acciones(self, record):
        return mark_safe(
            f'<a href="/dashboard/formularios/seguros/marchamo/{record["respuesta_id"]}/" '
            f'target="_blank" rel="noopener noreferrer" '
            f'class="text-primary hover:text-primary-dark transition-colors">'
            f'<span class="material-symbols-outlined">open_in_new</span>'
            f'</a>'
        )
 
    class Meta:
        template_name = "dashboard/components/table.html"
        attrs     = {"class": "w-full text-left border-collapse"}
        row_attrs = {"class": "hover:bg-surface-container-low transition-colors cursor-pointer"}
        sequence  = ("FechaHora", "Cedula", "NombreCompleto",
                     "TipoVehiculo", "NumeroPlaca", "SucursalRetiro", "acciones")
        


class AgentesTable(tables.Table):
    unidad_nombre = tables.Column(verbose_name="Unidad")
    sucursal_nombre = tables.Column(verbose_name="Sucursal")
    
    qr = tables.Column(
        verbose_name="QR",
        orderable=False,
        empty_values=(),
    )
    acciones = tables.Column(
        verbose_name="Acciones",
        orderable=False,
        empty_values=(),
    )

    class Meta:
        template_name = "dashboard/components/table.html"
        attrs = {"class": "w-full text-left border-collapse"}
        sequence = ("agente_id", "nombre", "unidad_nombre", "sucursal_nombre", "qr", "acciones")
        fields = ("agente_id", "nombre", "unidad_nombre", "sucursal_nombre")



