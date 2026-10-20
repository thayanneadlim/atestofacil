from django.contrib import admin
from django.urls import include, path
from django.conf import settings
from django.conf.urls.static import static
from django.contrib.auth import views as auth_views

urlpatterns = [
    path("admin/", admin.site.urls),
    path("accounts/login/", auth_views.LoginView.as_view(), name="login"),
    path(
        "accounts/logout/",
        auth_views.LogoutView.as_view(),
        name="logout",
    ),
    path("", include("core.urls")),
    path("pessoas/", include("pessoas.urls")),
    path("empresas/", include("empresas.urls")),
    path("textos-padroes/", include("textos_padroes.urls")),
    path("servicos/", include("servicos.urls")),
    path("documentos/", include("documentos.urls")),
    path("api/pessoas/", include("pessoas.api.urls")),
    path("api/empresas/", include("empresas.api.urls")),
    path("api/textos-padroes/", include("textos_padroes.api.urls")),
    path("api/servicos/", include("servicos.api.urls")),
    path("api/documentos/", include("documentos.api.urls")),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
