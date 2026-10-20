from rest_framework.routers import DefaultRouter

from pessoas.api.views import PessoaViewSet

router = DefaultRouter()
router.register("", PessoaViewSet, basename="api-pessoas")

urlpatterns = router.urls
