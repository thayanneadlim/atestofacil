from django.contrib import admin

from core.models import SequenciaAnual


@admin.register(SequenciaAnual)
class SequenciaAnualAdmin(admin.ModelAdmin):
    list_display = ("tipo", "ano", "ultimo_numero")
