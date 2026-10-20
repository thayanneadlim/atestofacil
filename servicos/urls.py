from django.urls import path
from servicos import views

app_name = "servicos"

urlpatterns = [
    path("", views.servico_list, name="list"),
    path("novo/", views.servico_create, name="create"),
    path("busca-cadastro/", views.busca_cadastro, name="busca_cadastro"),
    path("<int:pk>/", views.servico_detail, name="detail"),
    path("<int:pk>/editar/", views.servico_edit, name="edit"),
    path("<int:pk>/excel/", views.servico_excel, name="excel"),
    path("<int:pk>/excel-rpa/", views.servico_excel_rpa, name="excel_rpa"),
    path("<int:pk>/atesto/", views.servico_atesto, name="atesto"),
    path("<int:pk>/atesto-assinado/", views.servico_anexar_atesto, name="anexar_atesto"),
]
