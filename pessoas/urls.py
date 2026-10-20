from django.urls import path
from pessoas import views

app_name = "pessoas"

urlpatterns = [
    path("", views.pessoa_list, name="list"),
    path("nova/", views.pessoa_create, name="create"),
    path("<int:pk>/", views.pessoa_detail, name="detail"),
    path("<int:pk>/editar/", views.pessoa_edit, name="edit"),
]
