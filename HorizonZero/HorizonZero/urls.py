from django.contrib import admin
from django.urls import include, path
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path("usuarios/", include("apps.usuarios.urls", namespace="usuarios")),
    path("dashboard/", include("apps.dashboard.urls",  namespace="dashboard")),
]+ static(settings.STATIC_URL, document_root=settings.BASE_DIR / "static")
