from django.urls import path
from textos_padroes import views

app_name = "textos_padroes"

urlpatterns = [
    path("", views.texto_list, name="list"),
    path("novo/", views.texto_create, name="create"),
    path("<int:pk>/", views.texto_detail, name="detail"),
    path("<int:pk>/editar/", views.texto_edit, name="edit"),
    path("<int:pk>/excluir/", views.texto_delete, name="delete"),
]
