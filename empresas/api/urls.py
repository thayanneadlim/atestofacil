from rest_framework.routers import DefaultRouter

from empresas.api.views import EmpresaViewSet

router = DefaultRouter()
router.register("", EmpresaViewSet, basename="api-empresas")

urlpatterns = router.urls
