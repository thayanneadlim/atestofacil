from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect, render

from core.services.templates_texto import VARIAVEIS_SUPORTADAS, renderizar_texto
from textos_padroes.forms import TextoPadraoForm
from textos_padroes.models import TextoPadrao


@login_required
def texto_list(request):
    return render(request, "textos_padroes/list.html", {"textos": TextoPadrao.objects.all()})


@login_required
def texto_create(request):
    form = TextoPadraoForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Texto padrão criado.")
        return redirect("textos_padroes:list")
    return render(
        request,
        "textos_padroes/form.html",
        {"form": form, "titulo": "Novo texto padrão", "variaveis": sorted(VARIAVEIS_SUPORTADAS)},
    )


@login_required
def texto_edit(request, pk):
    texto = get_object_or_404(TextoPadrao, pk=pk)
    form = TextoPadraoForm(request.POST or None, instance=texto)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Texto padrão atualizado.")
        return redirect("textos_padroes:detail", pk=texto.pk)
    return render(
        request,
        "textos_padroes/form.html",
        {"form": form, "titulo": "Editar texto padrão", "texto": texto, "variaveis": sorted(VARIAVEIS_SUPORTADAS)},
    )


@login_required
def texto_detail(request, pk):
    texto = get_object_or_404(TextoPadrao, pk=pk)
    exemplo = {
        "nome": "João da Silva",
        "cpf": "000.000.000-00",
        "chave_pix": "joao@email.com",
        "numero_documento": "001/2026",
        "data_documento": "15/09/2026",
        "empresa": "Empresa Exemplo",
        "dias_trabalhados": (
            "05 de Setembro até 08 de Setembro 05/09 e 07/09: 08:00h às 17:00h\n"
            "06/09 e 08/09: 08:00h às 10:00h."
        ),
        "motivo_contratacao": "Cobertura de férias",
        "funcao": "Assistente",
        "total_horas": "22",
        "valor_total_formatado": "R$ 440,00",
        "nome_curto": "João Silva",
        "data_vencimento": "17/09/2026",
    }
    previa = renderizar_texto(texto.conteudo, exemplo)
    return render(request, "textos_padroes/detail.html", {"texto": texto, "previa": previa})


@login_required
def texto_delete(request, pk):
    texto = get_object_or_404(TextoPadrao, pk=pk)
    if request.method == "POST":
        if texto.em_uso():
            messages.error(request, "Texto já utilizado em documentos. Desative em vez de excluir.")
            return redirect("textos_padroes:detail", pk=pk)
        texto.delete()
        messages.success(request, "Texto removido.")
        return redirect("textos_padroes:list")
    return render(request, "textos_padroes/confirm_delete.html", {"texto": texto})
