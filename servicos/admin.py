from django.contrib import admin
from servicos.models import DiaTrabalhado, Servico


class DiaTrabalhadoInline(admin.TabularInline):
    model = DiaTrabalhado
    extra = 0


@admin.register(Servico)
class ServicoAdmin(admin.ModelAdmin):
    list_display = ("numero_documento", "numero_rpa_documento", "pessoa", "empresa", "data_servico")
    search_fields = ("numero_documento", "pessoa__nome", "empresa__nome")
    inlines = [DiaTrabalhadoInline]


@admin.register(DiaTrabalhado)
class DiaTrabalhadoAdmin(admin.ModelAdmin):
    list_display = ("servico", "data", "hora_inicio", "hora_fim", "quantidade_horas")
