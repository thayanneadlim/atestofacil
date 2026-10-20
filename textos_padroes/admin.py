from django.contrib import admin
from textos_padroes.models import TextoPadrao


@admin.register(TextoPadrao)
class TextoPadraoAdmin(admin.ModelAdmin):
    list_display = ("nome", "tipo_documento", "ativo", "updated_at")
    list_filter = ("tipo_documento", "ativo")
    search_fields = ("nome", "descricao")
