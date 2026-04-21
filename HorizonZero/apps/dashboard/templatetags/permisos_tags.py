from django import template
register = template.Library()

@register.filter
def tiene_perm(user, codename):
    return user.has_perm(f"dashboard.{codename}")



# Uso en template:
"""
{% load permisos_tags %}
 
{% if request.user|tiene_perm:"view_formulario_tarjetas" %}
  <a href="{% url 'dashboard:formulario_tarjetas' %}">Tarjetas</a>
{% endif %}
"""