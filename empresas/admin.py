from django.contrib import admin
from empresas.models import Empresa


@admin.register(Empresa)
class EmpresaAdmin(admin.ModelAdmin):
    list_display = ("nome", "cnpj", "ativo", "created_at")
    search_fields = ("nome", "cnpj")
    list_filter = ("ativo",)
