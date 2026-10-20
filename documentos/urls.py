from django.urls import path

from documentos import views

app_name = "documentos"

urlpatterns = [
    path("", views.documento_list, name="list"),
    path("<int:pk>/", views.documento_detail, name="detail"),
    path("<int:pk>/pdf/", views.documento_pdf, name="pdf"),
    path("<int:pk>/imprimir/", views.documento_print, name="print"),
    path("<int:pk>/versoes/<int:versao>/", views.documento_versao, name="versao"),
]
