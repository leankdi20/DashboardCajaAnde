from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth import get_user_model
from .models import PerfilUsuario

User = get_user_model()


class PerfilUsuarioInline(admin.StackedInline):
    model          = PerfilUsuario
    can_delete     = False
    verbose_name   = "Perfil y Restricciones"
    fields         = ("unidad", "sucursal")
    extra          = 1  # crea el perfil si no existe


class UserAdmin(BaseUserAdmin):
    inlines = (PerfilUsuarioInline,)


admin.site.unregister(User)
admin.site.register(User, UserAdmin)