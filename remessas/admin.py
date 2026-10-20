from django.contrib import admin
from remessas.models import Remessa


@admin.register(Remessa)
class RemessaAdmin(admin.ModelAdmin):
    list_display = ("numero_formatado", "pessoa", "status", "data_criacao")
    list_filter = ("status", "ano")
    search_fields = ("numero_formatado", "pessoa__nome", "pessoa__cpf")
