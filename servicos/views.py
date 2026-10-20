import json
import re
import unicodedata
from datetime import datetime

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import FileResponse, Http404, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_POST

from core.services.excel_template import gerar_excel_previsao, gerar_excel_rpa
from core.services.templates_texto import formatar_datas_trabalhadas
from core.services.word_template import gerar_documento_atesto
from core.services.servico_documentos import (
    ServicoDocumentoError,
    criar_servico_com_documentos,
    editar_servico_com_documentos,
    montar_contexto_de_servico,
)
from empresas.models import Empresa
from pessoas.models import Pessoa
from servicos.forms import AtestoAssinadoForm, ServicoForm
from servicos.models import Servico


def _parse_dias(dias_json: str):
    dados = json.loads(dias_json or "[]")
    dias = []
    for item in dados:
        data = datetime.strptime(item["data"], "%Y-%m-%d").date()
        hora_inicio = item.get("hora_inicio") or ""
        hora_fim = item.get("hora_fim") or ""
        if not hora_inicio or not hora_fim:
            raise ValueError("Informe hora de início e fim para cada dia.")
        dias.append(
            {
                "data": data,
                "hora_inicio": hora_inicio,
                "hora_fim": hora_fim,
                "observacao": item.get("observacao") or "",
            }
        )
    datas = [d["data"] for d in dias]
    if len(datas) != len(set(datas)):
        raise ValueError("Dia duplicado no serviço.")
    return dias


def normalizar_busca(texto: str) -> str:
    """Ignora acento, maiúsculas e pontuação para comparar nomes."""
    base = unicodedata.normalize("NFD", texto or "")
    sem_acento = "".join(ch for ch in base if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", sem_acento.casefold()).strip()


def _textos_selecionados(cleaned_data: dict) -> dict:
    textos = {"DOCUMENTO_1": cleaned_data["texto_documento_1"]}
    if cleaned_data.get("texto_documento_2"):
        textos["DOCUMENTO_2"] = cleaned_data["texto_documento_2"]
    if cleaned_data.get("texto_documento_3"):
        textos["DOCUMENTO_3"] = cleaned_data["texto_documento_3"]
    return textos


@login_required
def busca_cadastro(request):
    """Sugestões de empresa ou pessoa já cadastradas, para o autocomplete."""
    tipo = request.GET.get("tipo") or ""
    termo = normalizar_busca(request.GET.get("q") or "")
    if tipo == "empresa":
        registros = Empresa.objects.filter(ativo=True).order_by("nome")
    elif tipo == "pessoa":
        registros = Pessoa.objects.filter(ativo=True).order_by("nome")
    else:
        return JsonResponse({"results": []})
    if not termo:
        return JsonResponse({"results": []})
    resultados = []
    for registro in registros:
        if termo in normalizar_busca(registro.nome):
            resultados.append({"id": registro.pk, "nome": registro.nome})
            if len(resultados) >= 20:
                break
    return JsonResponse({"results": resultados})


@login_required
def servico_list(request):
    return render(
        request,
        "servicos/list.html",
        {"servicos": Servico.objects.select_related("pessoa", "empresa")},
    )


@login_required
def servico_create(request):
    form = ServicoForm(request.POST or None, initial={"dias_json": "[]"})
    if request.method == "POST" and form.is_valid():
        try:
            dias = _parse_dias(form.cleaned_data["dias_json"])
            textos = _textos_selecionados(form.cleaned_data)
            servico = criar_servico_com_documentos(
                pessoa=form.cleaned_data["pessoa"],
                empresa=form.cleaned_data["empresa"],
                motivo_contratacao=form.cleaned_data["motivo_contratacao"],
                funcao=form.cleaned_data["funcao"],
                data_servico=form.cleaned_data["data_servico"],
                assinatura=form.cleaned_data["assinatura"],
                dias_trabalhados=dias,
                textos_por_tipo=textos,
                usuario=request.user,
            )
            messages.success(
                request,
                f"Serviço {servico.numero_documento} criado com {servico.documentos.count()} documento(s).",
            )
            return redirect("servicos:detail", pk=servico.pk)
        except (ServicoDocumentoError, ValueError, json.JSONDecodeError) as exc:
            messages.error(request, str(exc))
    return render(
        request,
        "servicos/form.html",
        {"form": form, "titulo": "Novo serviço"},
    )


@login_required
def servico_edit(request, pk):
    servico = get_object_or_404(Servico.objects.prefetch_related("documentos"), pk=pk)
    dias_inicial = json.dumps(
        [
            {
                "data": d.data.isoformat(),
                "hora_inicio": d.hora_inicio.strftime("%H:%M"),
                "hora_fim": d.hora_fim.strftime("%H:%M"),
                "observacao": d.observacao,
            }
            for d in servico.dias_trabalhados.all()
        ]
    )
    initial = {"dias_json": dias_inicial}
    docs = {d.tipo_documento: d.texto_padrao_id for d in servico.documentos.all()}
    for tipo, field in (
        ("DOCUMENTO_1", "texto_documento_1"),
        ("DOCUMENTO_2", "texto_documento_2"),
        ("DOCUMENTO_3", "texto_documento_3"),
    ):
        if docs.get(tipo):
            initial[field] = docs[tipo]

    form = ServicoForm(request.POST or None, instance=servico, initial=initial)
    form.fields["pessoa"].disabled = True
    if request.method == "POST" and form.is_valid():
        try:
            dias = _parse_dias(form.cleaned_data["dias_json"])
            editar_servico_com_documentos(
                servico=servico,
                motivo_contratacao=form.cleaned_data["motivo_contratacao"],
                funcao=form.cleaned_data["funcao"],
                data_servico=form.cleaned_data["data_servico"],
                assinatura=form.cleaned_data["assinatura"],
                dias_trabalhados=dias,
                textos_por_tipo=_textos_selecionados(form.cleaned_data),
                usuario=request.user,
                motivo_alteracao=form.cleaned_data.get("motivo_alteracao")
                or "Edição do serviço",
            )
            messages.success(request, "Serviço atualizado e documentos regenerados.")
            return redirect("servicos:detail", pk=servico.pk)
        except (ServicoDocumentoError, ValueError, json.JSONDecodeError) as exc:
            messages.error(request, str(exc))
    return render(
        request,
        "servicos/form.html",
        {
            "form": form,
            "titulo": f"Editar serviço {servico.numero_documento}",
            "servico": servico,
        },
    )


@login_required
def servico_detail(request, pk):
    servico = get_object_or_404(
        Servico.objects.select_related("pessoa", "empresa").prefetch_related(
            "documentos", "dias_trabalhados"
        ),
        pk=pk,
    )
    return render(
        request,
        "servicos/detail.html",
        {"servico": servico, "resumo": servico.resumo_pagamento},
    )


def _voltar_depois_do_anexo(request, servico):
    destino = request.POST.get("next") or ""
    if url_has_allowed_host_and_scheme(
        destino,
        allowed_hosts={request.get_host()},
        require_https=request.is_secure(),
    ):
        return destino
    return reverse("servicos:detail", args=[servico.pk])


@login_required
@require_POST
def servico_anexar_atesto(request, pk):
    servico = get_object_or_404(Servico, pk=pk)
    form = AtestoAssinadoForm(request.POST, request.FILES, instance=servico)
    if form.is_valid():
        form.save()
        messages.success(request, "Atesto assinado anexado.")
    else:
        erros = form.errors.get("atesto_assinado")
        messages.error(
            request,
            erros[0] if erros else "Não foi possível anexar o atesto assinado.",
        )
    return redirect(_voltar_depois_do_anexo(request, servico))


@login_required
def servico_excel(request, pk):
    servico = get_object_or_404(
        Servico.objects.select_related("pessoa", "empresa").prefetch_related(
            "dias_trabalhados", "documentos"
        ),
        pk=pk,
    )
    try:
        contexto = montar_contexto_de_servico(servico)
        arquivo = gerar_excel_previsao(
            contexto=contexto,
            numero_documento=servico.numero_documento,
        )
    except FileNotFoundError as exc:
        raise Http404(str(exc)) from exc

    return FileResponse(
        arquivo,
        as_attachment=True,
        filename=arquivo.name,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@login_required
def servico_excel_rpa(request, pk):
    servico = get_object_or_404(
        Servico.objects.select_related("pessoa", "empresa").prefetch_related(
            "dias_trabalhados", "documentos"
        ),
        pk=pk,
    )
    try:
        contexto = montar_contexto_de_servico(servico)
        arquivo = gerar_excel_rpa(
            contexto=contexto,
            numero_documento=servico.numero_rpa_documento,
        )
    except FileNotFoundError as exc:
        raise Http404(str(exc)) from exc

    return FileResponse(
        arquivo,
        as_attachment=True,
        filename=arquivo.name,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@login_required
def servico_atesto(request, pk):
    servico = get_object_or_404(
        Servico.objects.select_related("pessoa", "empresa").prefetch_related(
            "dias_trabalhados", "documentos"
        ),
        pk=pk,
    )
    try:
        contexto = montar_contexto_de_servico(servico)
        contexto["dias_trabalhados"] = formatar_datas_trabalhadas(
            servico.dias_trabalhados.all()
        )
        arquivo = gerar_documento_atesto(
            contexto=contexto,
            numero_documento=servico.numero_documento,
        )
    except FileNotFoundError as exc:
        raise Http404(str(exc)) from exc

    return FileResponse(
        arquivo,
        as_attachment=True,
        filename=arquivo.name,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
