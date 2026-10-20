from django.urls import path

from core import views

app_name = "core"

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("numeracao/", views.numeracao, name="numeracao"),
]
