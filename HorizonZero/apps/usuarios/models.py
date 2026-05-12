from django.db import models
from django.contrib.auth import get_user_model

User = get_user_model()


class PerfilUsuario(models.Model):
    """
    Perfil extendido del usuario.
    Permite restringir el acceso a datos por unidad y/o sucursal.
    Si unidad/sucursal están vacíos → el usuario ve todo (sin restricción).
    Si tienen valor → el usuario solo ve datos de esa unidad/sucursal.
    """
    usuario  = models.OneToOneField(
        User,
        on_delete=models.CASCADE,
        related_name="perfil",
    )
    unidad   = models.CharField(
        max_length=100,
        blank=True,
        null=True,
        verbose_name="Unidad restringida",
        help_text="Si se asigna, el usuario solo verá encuestas de esta unidad.",
    )
    sucursal = models.CharField(
        max_length=100,
        blank=True,
        null=True,
        verbose_name="Sucursal restringida",
        help_text="Si se asigna, el usuario solo verá encuestas de esta sucursal.",
    )

    class Meta:
        verbose_name        = "Perfil de Usuario"
        verbose_name_plural = "Perfiles de Usuario"

    def __str__(self):
        partes = [self.usuario.username]
        if self.unidad:
            partes.append(f"Unidad: {self.unidad}")
        if self.sucursal:
            partes.append(f"Sucursal: {self.sucursal}")
        return " — ".join(partes)

    def tiene_restriccion(self) -> bool:
        """Retorna True si el usuario tiene alguna restricción activa."""
        return bool(self.unidad or self.sucursal)