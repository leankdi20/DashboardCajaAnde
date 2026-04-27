from django.contrib import admin
from django.urls import include, path, re_path
from django.conf import settings
from django.conf.urls.static import static
from django.views.static import serve
from django.http import HttpResponse
urlpatterns = [
    path('admin/', admin.site.urls),
    path("__test_urls__/", lambda request: HttpResponse("OK URLS")),
    path("usuarios/", include("apps.usuarios.urls", namespace="usuarios")),
    path("dashboard/", include("apps.dashboard.urls", namespace="dashboard")),
    
]

if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=str(settings.BASE_DIR / "static"))
    urlpatterns += static(settings.MEDIA_URL, document_root=str(settings.MEDIA_ROOT))