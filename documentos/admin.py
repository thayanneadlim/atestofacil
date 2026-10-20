from django.contrib import admin

from documentos.models import Documento, DocumentoVersao


class DocumentoVersaoInline(admin.TabularInline):
    model = DocumentoVersao
    extra = 0
    readonly_fields = ("versao", "motivo_alteracao", "created_at", "criado_por")


@admin.register(Documento)
class DocumentoAdmin(admin.ModelAdmin):
    list_display = (
        "numero_documento",
        "tipo_documento",
        "pessoa",
        "versao_atual",
        "data_documento",
    )
    list_filter = ("tipo_documento",)
    search_fields = ("numero_documento", "pessoa__nome")
    inlines = [DocumentoVersaoInline]


@admin.register(DocumentoVersao)
class DocumentoVersaoAdmin(admin.ModelAdmin):
    list_display = ("documento", "versao", "motivo_alteracao", "created_at")
