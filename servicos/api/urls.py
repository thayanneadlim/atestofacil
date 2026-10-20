from rest_framework.routers import DefaultRouter

from servicos.api.views import ServicoViewSet

router = DefaultRouter()
router.register("", ServicoViewSet, basename="api-servicos")

urlpatterns = router.urls
