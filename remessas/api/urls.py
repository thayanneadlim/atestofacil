from rest_framework.routers import DefaultRouter

from remessas.api.views import RemessaViewSet

router = DefaultRouter()
router.register("", RemessaViewSet, basename="api-remessas")

urlpatterns = router.urls
