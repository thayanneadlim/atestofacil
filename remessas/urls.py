from django.urls import path
from remessas import views

app_name = "remessas"

urlpatterns = [
    path("", views.remessa_list, name="list"),
    path("criar/", views.remessa_create, name="create"),
    path("<int:pk>/", views.remessa_detail, name="detail"),
    path("<int:pk>/editar/", views.remessa_edit, name="edit"),
    path("<int:pk>/cancelar/", views.remessa_cancel, name="cancel"),
]
