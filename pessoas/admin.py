from django.contrib import admin
from pessoas.models import Pessoa


@admin.register(Pessoa)
class PessoaAdmin(admin.ModelAdmin):
    list_display = ("nome", "cpf", "ativo", "created_at")
    search_fields = ("nome", "cpf")
    list_filter = ("ativo",)
