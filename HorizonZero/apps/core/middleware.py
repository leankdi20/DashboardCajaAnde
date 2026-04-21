class LimpiarCachePermisosMiddleware:
    """
    Limpia el caché de permisos de Django en cada request.
    Necesario cuando se asignan grupos con sesión activa.
    """
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated:
            for attr in ('_perm_cache', '_user_perm_cache', '_group_perm_cache'):
                try:
                    delattr(request.user, attr)
                except AttributeError:
                    pass
        return self.get_response(request)