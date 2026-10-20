from rest_framework.routers import DefaultRouter

from textos_padroes.api.views import TextoPadraoViewSet

router = DefaultRouter()
router.register("", TextoPadraoViewSet, basename="api-textos-padroes")

urlpatterns = router.urls
