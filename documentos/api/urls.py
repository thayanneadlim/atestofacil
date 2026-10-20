from rest_framework.routers import DefaultRouter

from documentos.api.views import DocumentoViewSet

router = DefaultRouter()
router.register("", DocumentoViewSet, basename="api-documentos")

urlpatterns = router.urls
