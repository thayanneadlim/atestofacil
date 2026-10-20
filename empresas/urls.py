from django.urls import path
from empresas import views

app_name = "empresas"

urlpatterns = [
    path("", views.empresa_list, name="list"),
    path("nova/", views.empresa_create, name="create"),
    path("<int:pk>/", views.empresa_detail, name="detail"),
    path("<int:pk>/editar/", views.empresa_edit, name="edit"),
    path("<int:pk>/excluir/", views.empresa_delete, name="delete"),
]
