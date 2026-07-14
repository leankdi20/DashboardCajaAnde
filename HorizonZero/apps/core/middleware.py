from django.http import Http404
from django.shortcuts import render


class LimpiarCachePermisosMiddleware:
    """
    Limpia el caché de permisos de Django en cada request.
    Necesario cuando se asignan grupos con sesión activa.
    """
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated:
            for attr in ("_perm_cache", "_user_perm_cache", "_group_perm_cache"):
                try:
                    delattr(request.user, attr)
                except AttributeError:
                    pass
        return self.get_response(request)


class FriendlyNotFoundMiddleware:
    """
    Muestra la pagina 404 amigable incluso cuando DEBUG=True.
    """

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        try:
            response = self.get_response(request)
        except Http404:
            return render(request, "errors/404.html", status=404)

        if response.status_code == 404 and "text/html" in response.get("Content-Type", ""):
            return render(request, "errors/404.html", status=404)

        return response
