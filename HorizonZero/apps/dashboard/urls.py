from django.urls import path
from . import views

app_name = "dashboard"




urlpatterns_agentes = [
    path("usuarios/agentes/",                          views.agentes_home,           name="agentes_home"),
    path("usuarios/agentes/crear/",                    views.agente_crear,            name="agente_crear"),
    path("usuarios/agentes/<int:agente_id>/editar/",   views.agente_editar,           name="agente_editar"),
    path("usuarios/agentes/<int:agente_id>/eliminar/", views.agente_eliminar,         name="agente_eliminar"),
    path("usuarios/agentes/<int:agente_id>/qr/",       views.agente_qr,               name="agente_qr"),
    path("usuarios/agentes/<int:agente_id>/qr/<int:encuesta_id>/descargar/",
                                                       views.agente_qr_download,      name="agente_qr_download"),
    path("usuarios/agentes/<int:agente_id>/qr/zip/",   views.agente_qr_download_zip,  name="agente_qr_download_zip"),


    path("usuarios/agentes/inactivos/",                        views.agentes_inactivos, name="agentes_inactivos"),
    path("usuarios/agentes/<int:agente_id>/restaurar/",        views.agente_restaurar,  name="agente_restaurar"),


    
]



urlpatterns_logs = [
    path("logs/",                 views.logs_home,         name="logs_home"),
    path("logs/<int:log_id>/",    views.log_detalle,       name="log_detalle"),
    path("logs/export/excel/",    views.logs_export_excel, name="logs_export_excel"),
    path("logs/export/pdf/",      views.logs_export_pdf,   name="logs_export_pdf"),
]

urlpatterns = [
    path("", views.dashboard_home, name="home"),

    # Encuestas SATISFACCION
    path("encuestas/satisfaccion/",        views.encuesta_satisfaccion,         name="encuesta_satisfaccion"),
    path("encuestas/satisfaccion/<int:respuesta_id>/", 
        views.encuesta_satisfaccion_detalle, name="encuesta_satisfaccion_detalle"),
    path("encuestas/satisfaccion/<int:respuesta_id>/exportar/", 
        views.encuesta_satisfaccion_detalle_exportar, name="encuesta_satisfaccion_detalle_exportar"),
    path("encuestas/satisfaccion/exportar/", views.encuesta_satisfaccion_exportar, name="encuesta_satisfaccion_exportar"),



    # OFICINA DIGITAL
    path("encuestas/satisfaccion_of_dig/", 
            views.encuesta_satisfaccion_oficina, 
            name="encuesta_satisfaccion_oficina"),
    path("encuestas/satisfaccion_of_dig/<int:respuesta_id>/", 
            views.encuesta_satisfaccion_detalle_of_dig, 
            name="encuesta_satisfaccion_oficina_detalle"),
    path("encuestas/satisfaccion_of_dig/<int:respuesta_id>/exportar/", 
            views.encuesta_satisfaccion_detalle_of_dig_exportar, 
            name="encuesta_satisfaccion_oficina_detalle_exportar"),
    path("encuestas/satisfaccion_of_dig/exportar/",     
            views.encuesta_satisfaccion_of_dig_exportar, 
            name="encuesta_satisfaccion_oficina_exportar"),



    # EXPERIENCIA EN PÁGINA WEB
    path("encuestas/satisfaccion_experiencia_web/",     views.encuesta_experiencia_web,      name="encuesta_experiencia_web"),
    path("encuestas/experiencia-web/<int:respuesta_id>/",
        views.encuesta_experiencia_web_detalle,
        name="encuesta_experiencia_web_detalle"),
    path("encuestas/experiencia-web/exportar/",
        views.encuesta_experiencia_web_exportar,
        name="encuesta_experiencia_web_exportar"),



    # ENCUESTA SATISFACCION WHATSAPP AGENTE
    path("encuestas/whatsapp-agente/",     views.encuesta_whatsApp_agente,      name="encuesta_whatsapp_agente"),
    path("encuestas/whatsapp-agente/<int:respuesta_id>/",
            views.encuesta_whatsapp_agente_detalle,
            name="encuesta_whatsapp_agente_detalle"),
    path("encuestas/whatsapp-agente/exportar/",
            views.encuesta_whatsapp_agente_exportar,
            name="encuesta_whatsapp_agente_exportar"),
    path("encuestas/whatsapp-agente/<int:respuesta_id>/exportar",
            views.encuesta_whatsapp_agente_detalle_exportar,
            name="encuesta_whatsapp_agente_detalle_exportar"),



    # ENCUESTA SATISFACCION WHATSAPP
    path("encuestas/whatsapp/", views.encuesta_whatsApp_,  name="encuesta_whatsapp"),
    path("encuestas/whatsapp/<int:respuesta_id>/",
            views.encuesta_whatsapp_detalle,
            name="encuesta_whatsapp_detalle"),
    path("encuestas/whatsapp/exportar/",
            views.encuesta_whatsapp_exportar,
            name="encuesta_whatsapp_exportar"),
    path("encuestas/whatsapp/<int:respuesta_id>/exportar",
            views.encuesta_whatsapp_detalle_exportar,
            name="encuesta_whatsapp_detalle_exportar"),


    path("encuestas/feria-salud/",      
         views.encuesta_feria_salud_,        
         name="encuesta_feria_salud"),
#     path("encuestas/feria-salud/<int:respuesta_id>/",
#         views.encuesta_feria_salud_detalle,
#         name="encuesta_feria_salud_detalle"),
    path("encuestas/feria-salud/exportar/",
        views.encuesta_feria_salud_exportar,
        name="encuesta_feria_salud_exportar"),







    # Formularios Tarjetas
    path("formularios/tarjetas/",              views.formulario_tarjetas,              name="formulario_tarjetas"),

        #Tarjeta de crédito
        path(
        "formularios/tarjetas/solicitudes/buscar/<str:campo>/",
        views.soli_tarj_credito_buscar,
        name="soli_tarj_credito_buscar",
        ),

        path("formularios/tarjetas/solicitudes/",
        views.soli_tarj_credito_lista,
        name="soli_tarj_credito_lista",
        ),
        path("formularios/tarjetas/solicitudes/<int:respuesta_id>/",
        views.soli_tarj_credito_detalle,
        name="soli_tarj_credito_detalle",
        ),
        path("formularios/tarjetas/solicitudes/export/",
        views.soli_tarj_credito_export,
        name="soli_tarj_credito_export",
        ),
        path(
        "formularios/tarjetas/solicitudes/<int:respuesta_id>/export/",
        views.soli_tarj_credito_export_detalle,
        name="soli_tarj_credito_export_detalle",
        ),
 
   
   
   
         # fORMULARIO Tarjeta de debito Ciudadano Oro

        path("formularios/tarjetas/debito/buscar/<str:campo>/",
        views.soli_tarj_debito_buscar,
        name="soli_tarj_debito_buscar"),
        path("formularios/tarjetas/debito/",
        views.soli_tarj_debito_lista,
        name="soli_tarj_debito_lista"),
        path("formularios/tarjetas/debito/<int:respuesta_id>/",
        views.soli_tarj_debito_detalle,
        name="soli_tarj_debito_detalle"),
        path("formularios/tarjetas/debito/export/",
        views.soli_tarj_debito_export,
        name="soli_tarj_debito_export"),




        #Formulario Tarjeta de debito común
        path("formularios/tarjetas/debito-gestion/buscar/<str:campo>/",
        views.soli_tarj_debito_gestion_buscar,
        name="soli_tarj_debito_gestion_buscar"),
        path("formularios/tarjetas/debito-gestion/",
        views.soli_tarj_debito_gestion_lista,
        name="soli_tarj_debito_gestion_lista"),
        path("formularios/tarjetas/debito-gestion/<int:respuesta_id>/",
        views.soli_tarj_debito_gestion_detalle,
        name="soli_tarj_debito_gestion_detalle"),
        path("formularios/tarjetas/debito-gestion/export/",
        views.soli_tarj_debito_gestion_export,
        name="soli_tarj_debito_gestion_export"),



        #Formulario Tarjeta Solicitud para redención de puntos

        path("formularios/tarjetas/redencion-puntos/buscar/<str:campo>/",
        views.soli_redencion_puntos_buscar,
        name="soli_redencion_puntos_buscar"),
        path("formularios/tarjetas/redencion-puntos/",
        views.soli_redencion_puntos_lista,
        name="soli_redencion_puntos_lista"),
        path("formularios/tarjetas/redencion-puntos/<int:respuesta_id>/",
        views.soli_redencion_puntos_detalle,
        name="soli_redencion_puntos_detalle"),
        path("formularios/tarjetas/redencion-puntos/export/",
        views.soli_redencion_puntos_export,
        name="soli_redencion_puntos_export"),
        path(
        "formularios/tarjetas/redencion-puntos/<int:respuesta_id>/export/",
        views.soli_redencion_puntos_export_detalle,
        name="soli_redencion_puntos_export_detalle",
        ),

        # Formulario Caja de ANDE Asistencia
        path("formularios/tarjetas/asistencia/buscar/<str:campo>/",
        views.caja_ande_asistencia_buscar,
        name="caja_ande_asistencia_buscar"),
        path("formularios/tarjetas/asistencia/",
        views.caja_ande_asistencia_lista,
        name="caja_ande_asistencia_lista"),
        path("formularios/tarjetas/asistencia/<int:respuesta_id>/",
        views.caja_ande_asistencia_detalle,
        name="caja_ande_asistencia_detalle"),
        path("formularios/tarjetas/asistencia/<int:respuesta_id>/export/",
        views.caja_ande_asistencia_export_detalle,
        name="caja_ande_asistencia_export_detalle"),
        path("formularios/tarjetas/asistencia/export/",
        views.caja_ande_asistencia_export,
        name="caja_ande_asistencia_export"),



        # Formularios Ahorros
    path("formularios/ahorros/",               views.formulario_ahorros,               name="formulario_ahorros"),

        #Formulario Solicitud Deposito Salario
        path("formularios/ahorros/deposito-salario/buscar/cedula/",
        views.soli_deposito_salario_buscar_cedula,
        name="soli_deposito_salario_buscar_cedula"),
        path("formularios/ahorros/deposito-salario/",
        views.soli_deposito_salario_lista,
        name="soli_deposito_salario_lista"),
        path("formularios/ahorros/deposito-salario/<int:respuesta_id>/",
        views.soli_deposito_salario_detalle,
        name="soli_deposito_salario_detalle"),
        path("formularios/ahorros/deposito-salario/<int:respuesta_id>/export/",
        views.soli_deposito_salario_export_detalle,
        name="soli_deposito_salario_export_detalle"),
        path("formularios/ahorros/deposito-salario/export/",
        views.soli_deposito_salario_export,
        name="soli_deposito_salario_export"),



        #Formulario Solicitud de ahorro: Modificación de cuota
        path("formularios/ahorros/modificacion-cuota/buscar/<str:campo>/",
        views.soli_ahorro_mod_cuota_buscar,
        name="soli_ahorro_mod_cuota_buscar"),
        path("formularios/ahorros/modificacion-cuota/",
        views.soli_ahorro_mod_cuota_lista,
        name="soli_ahorro_mod_cuota_lista"),
        path("formularios/ahorros/modificacion-cuota/<int:respuesta_id>/",
        views.soli_ahorro_mod_cuota_detalle,
        name="soli_ahorro_mod_cuota_detalle"),
        path("formularios/ahorros/modificacion-cuota/<int:respuesta_id>/export/",
        views.soli_ahorro_mod_cuota_export_detalle,
        name="soli_ahorro_mod_cuota_export_detalle"),
        path("formularios/ahorros/modificacion-cuota/export/",
        views.soli_ahorro_mod_cuota_export,
        name="soli_ahorro_mod_cuota_export"),


        #Formulario Solicitud de reinversion de ahorro
        path("formularios/ahorros/reinversion/buscar/<str:campo>/",
        views.soli_reinversion_ahorro_buscar,
        name="soli_reinversion_ahorro_buscar"),
        path("formularios/ahorros/reinversion/",
        views.soli_reinversion_ahorro_lista,
        name="soli_reinversion_ahorro_lista"),
        path("formularios/ahorros/reinversion/<int:respuesta_id>/",
        views.soli_reinversion_ahorro_detalle,
        name="soli_reinversion_ahorro_detalle"),
        path("formularios/ahorros/reinversion/<int:respuesta_id>/export/",
        views.soli_reinversion_ahorro_export_detalle,
        name="soli_reinversion_ahorro_export_detalle"),
        path("formularios/ahorros/reinversion/export/",
        views.soli_reinversion_ahorro_export,
        name="soli_reinversion_ahorro_export"),



        #Formulario Solicitud Autorizacion ahorro nuevo
        path("formularios/ahorros/autorizacion-nuevo/buscar/<str:campo>/",
        views.soli_autorizacion_ahorro_nuevo_buscar,
        name="soli_autorizacion_ahorro_nuevo_buscar"),
        path("formularios/ahorros/autorizacion-nuevo/",
        views.soli_autorizacion_ahorro_nuevo_lista,
        name="soli_autorizacion_ahorro_nuevo_lista"),
        path("formularios/ahorros/autorizacion-nuevo/<int:respuesta_id>/",
        views.soli_autorizacion_ahorro_nuevo_detalle,
        name="soli_autorizacion_ahorro_nuevo_detalle"),
        path("formularios/ahorros/autorizacion-nuevo/<int:respuesta_id>/export/",
        views.soli_autorizacion_ahorro_nuevo_export_detalle,
        name="soli_autorizacion_ahorro_nuevo_export_detalle"),
        path("formularios/ahorros/autorizacion-nuevo/export/",
        views.soli_autorizacion_ahorro_nuevo_export,
        name="soli_autorizacion_ahorro_nuevo_export"),





        # Formularios Vivienda
    path("formularios/vivienda/",              views.formulario_vivienda,              name="formulario_vivienda"),


    #Formulario Solicitud Prestamo de Vehiculo.
    path("formularios/vivienda/vehiculo/buscar/<str:campo>/",
        views.soli_compra_vehiculo_buscar,
        name="soli_compra_vehiculo_buscar"),
    path("formularios/vivienda/vehiculo/",
        views.soli_compra_vehiculo_lista,
        name="soli_compra_vehiculo_lista"),
    path("formularios/vivienda/vehiculo/<int:respuesta_id>/",
        views.soli_compra_vehiculo_detalle,
        name="soli_compra_vehiculo_detalle"),
    path("formularios/vivienda/vehiculo/<int:respuesta_id>/export/",
        views.soli_compra_vehiculo_export_detalle,
        name="soli_compra_vehiculo_export_detalle"),
    path("formularios/vivienda/vehiculo/export/",
        views.soli_compra_vehiculo_export,
        name="soli_compra_vehiculo_export"),





    #Formulario Vivienda Prestamo vivienda
    path("formularios/vivienda/prestamo/buscar/<str:campo>/",
        views.soli_prestamo_vivienda_buscar,
        name="soli_prestamo_vivienda_buscar"),
    path("formularios/vivienda/prestamo/",
        views.soli_prestamo_vivienda_lista,
        name="soli_prestamo_vivienda_lista"),
    path("formularios/vivienda/prestamo/<int:respuesta_id>/",
        views.soli_prestamo_vivienda_detalle,
        name="soli_prestamo_vivienda_detalle"),
    path("formularios/vivienda/prestamo/<int:respuesta_id>/export/",
        views.soli_prestamo_vivienda_export_detalle,
        name="soli_prestamo_vivienda_export_detalle"),
    path("formularios/vivienda/prestamo/export/",
        views.soli_prestamo_vivienda_export,
        name="soli_prestamo_vivienda_export"),



    #Formulario Desarrollo Economico
    path("formularios/vivienda/desarrollo/buscar/<str:campo>/",
        views.soli_prestamo_desarrollo_buscar,
        name="soli_prestamo_desarrollo_buscar"),
    path("formularios/vivienda/desarrollo/",
        views.soli_prestamo_desarrollo_lista,
        name="soli_prestamo_desarrollo_lista"),
    path("formularios/vivienda/desarrollo/<int:respuesta_id>/",
        views.soli_prestamo_desarrollo_detalle,
        name="soli_prestamo_desarrollo_detalle"),
    path("formularios/vivienda/desarrollo/<int:respuesta_id>/export/",
        views.soli_prestamo_desarrollo_export_detalle,
        name="soli_prestamo_desarrollo_export_detalle"),
    path("formularios/vivienda/desarrollo/export/",
        views.soli_prestamo_desarrollo_export,
        name="soli_prestamo_desarrollo_export"),


    #Formularios Préstamos
    path("formularios/prestamos/",             views.formulario_prestamos,             name="formulario_prestamos"),


    #Formulario Pre solicitud de credito personal
    path("formularios/prestamos/credito-personal/buscar/<str:campo>/",
        views.soli_presolicitud_credito_personal_buscar,
        name="soli_presolicitud_credito_personal_buscar"),
    path("formularios/prestamos/credito-personal/",
        views.soli_presolicitud_credito_personal_lista,
        name="soli_presolicitud_credito_personal_lista"),
    path("formularios/prestamos/credito-personal/<int:respuesta_id>/",
        views.soli_presolicitud_credito_personal_detalle,
        name="soli_presolicitud_credito_personal_detalle"),
    path("formularios/prestamos/credito-personal/<int:respuesta_id>/export/",
        views.soli_presolicitud_credito_personal_export_detalle,
        name="soli_presolicitud_credito_personal_export_detalle"),
    path("formularios/prestamos/credito-personal/export/",
        views.soli_presolicitud_credito_personal_export,
        name="soli_presolicitud_credito_personal_export"),

    # Formularios Control de crédito
    path("formularios/control-credito/",       views.formulario_control_credito,       name="formulario_control_credito"),

    # Formulario Comprobante Autorización Ahorro
    path("formularios/control-credito/autorizacion-ahorro/buscar/<str:campo>/",
        views.comprobante_autorizacion_ahorro_buscar,
        name="comprobante_autorizacion_ahorro_buscar"),
    path("formularios/control-credito/autorizacion-ahorro/",
        views.comprobante_autorizacion_ahorro_lista,
        name="comprobante_autorizacion_ahorro_lista"),
    path("formularios/control-credito/autorizacion-ahorro/<int:respuesta_id>/",
        views.comprobante_autorizacion_ahorro_detalle,
        name="comprobante_autorizacion_ahorro_detalle"),
    path("formularios/control-credito/autorizacion-ahorro/<int:respuesta_id>/export/",
        views.comprobante_autorizacion_ahorro_export_detalle,
        name="comprobante_autorizacion_ahorro_export_detalle"),
    path("formularios/control-credito/autorizacion-ahorro/export/",
        views.comprobante_autorizacion_ahorro_export,
        name="comprobante_autorizacion_ahorro_export"),
    
    # Formulario Comprobantes Pago
    path("formularios/control-credito/comprobantes-pago/buscar/<str:campo>/",
        views.comprobantes_pago_buscar,
        name="comprobantes_pago_buscar"),
    path("formularios/control-credito/comprobantes-pago/",
        views.comprobantes_pago_lista,
        name="comprobantes_pago_lista"),
    path("formularios/control-credito/comprobantes-pago/<int:respuesta_id>/",
        views.comprobantes_pago_detalle,
        name="comprobantes_pago_detalle"),
    path("formularios/control-credito/comprobantes-pago/<int:respuesta_id>/export/",
        views.comprobantes_pago_export_detalle,
        name="comprobantes_pago_export_detalle"),
    path("formularios/control-credito/comprobantes-pago/export/",
        views.comprobantes_pago_export,
        name="comprobantes_pago_export"),




    # Formulario Servicio Accionista
    path("formularios/servicio-accionista/",   views.formulario_servicio_accionista,   name="formulario_servicio_accionista"),

    # Formulario Caja tel
    path("formularios/servicio-accionista/clave-cajatel/buscar/<str:campo>/",
        views.soli_clave_temporal_cajatel_buscar,
        name="soli_clave_temporal_cajatel_buscar"),
    path("formularios/servicio-accionista/clave-cajatel/",
        views.soli_clave_temporal_cajatel_lista,
        name="soli_clave_temporal_cajatel_lista"),
    path("formularios/servicio-accionista/clave-cajatel/<int:respuesta_id>/",
        views.soli_clave_temporal_cajatel_detalle,
        name="soli_clave_temporal_cajatel_detalle"),
    path("formularios/servicio-accionista/clave-cajatel/<int:respuesta_id>/export/",
        views.soli_clave_temporal_cajatel_export_detalle,
        name="soli_clave_temporal_cajatel_export_detalle"),
    path("formularios/servicio-accionista/clave-cajatel/export/",
        views.soli_clave_temporal_cajatel_export,
        name="soli_clave_temporal_cajatel_export"),


    # Formularios Seguros
    path("formularios/seguros/",               views.formulario_seguros,               name="formulario_seguros"),



    # Formulario solicitudo seguro viajero
    path("formularios/seguros/viajero/buscar/<str:campo>/",
        views.soli_seguro_viajero_buscar,
        name="soli_seguro_viajero_buscar"),
    path("formularios/seguros/viajero/",
        views.soli_seguro_viajero_lista,
        name="soli_seguro_viajero_lista"),
    path("formularios/seguros/viajero/<int:respuesta_id>/",
        views.soli_seguro_viajero_detalle,
        name="soli_seguro_viajero_detalle"),
    path("formularios/seguros/viajero/<int:respuesta_id>/export/",
        views.soli_seguro_viajero_export_detalle,
        name="soli_seguro_viajero_export_detalle"),
    path("formularios/seguros/viajero/export/",
        views.soli_seguro_viajero_export,
        name="soli_seguro_viajero_export"),


    # Formulario Solicitud Marchamo
    path("formularios/seguros/marchamo/buscar/<str:campo>/",
     views.soli_marchamo_buscar,
     name="soli_marchamo_buscar"),
    path("formularios/seguros/marchamo/",
        views.soli_marchamo_lista,
        name="soli_marchamo_lista"),
    path("formularios/seguros/marchamo/<int:respuesta_id>/",
        views.soli_marchamo_detalle,
        name="soli_marchamo_detalle"),
    path("formularios/seguros/marchamo/<int:respuesta_id>/export/",
        views.soli_marchamo_export_detalle,
        name="soli_marchamo_export_detalle"),
    path("formularios/seguros/marchamo/export/",
        views.soli_marchamo_export,
        name="soli_marchamo_export"),

    *urlpatterns_agentes,
    *urlpatterns_logs,
]

 



"""
# ---

# Con esto listo, en el admin de Django se crean los grupos así:
# ```

Grupo "Administrador"
    ✅ Todos los view_*

Grupo "Encuestas Only"
    ✅ view_encuestas
    ✅ view_encuesta_satisfaccion
    ✅ view_encuesta_feria_salud
    ❌ view_formularios (no ve ese módulo)

Grupo "Formularios Tarjetas"
    ✅ view_formularios
    ✅ view_formulario_tarjetas
    ❌ view_formulario_ahorros

"""