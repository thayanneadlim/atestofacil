import json
from datetime import datetime
from decimal import Decimal

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.shortcuts import get_object_or_404, redirect, render

from core.services.remessa import RemessaError, criar_remessa_completa, editar_remessa
from remessas.forms import RemessaCreateForm, RemessaFilterForm
from remessas.models import Remessa


def _parse_dias(dias_json: str):
    dados = json.loads(dias_json or "[]")
    dias = []
    for item in dados:
        dias.append(
            {
                "data": datetime.strptime(item["data"], "%Y-%m-%d").date(),
                "quantidade_horas": Decimal(str(item["quantidade_horas"])),
                "observacao": item.get("observacao") or "",
            }
        )
    return dias


@login_required
def remessa_list(request):
    form = RemessaFilterForm(request.GET or None)
    qs = Remessa.objects.select_related("pessoa", "servico__empresa")
    if form.is_valid():
        q = form.cleaned_data.get("q") or ""
        if q:
            cpf = q.replace(".", "").replace("-", "")
            qs = qs.filter(
                Q(numero_formatado__icontains=q)
                | Q(pessoa__nome__icontains=q)
                | Q(pessoa__cpf__icontains=cpf)
            )
        empresa = form.cleaned_data.get("empresa")
        if empresa:
            qs = qs.filter(servico__empresa=empresa)
        if form.cleaned_data.get("data_inicio"):
            qs = qs.filter(data_criacao__gte=form.cleaned_data["data_inicio"])
        if form.cleaned_data.get("data_fim"):
            qs = qs.filter(data_criacao__lte=form.cleaned_data["data_fim"])
    return render(request, "remessas/list.html", {"remessas": qs, "form": form})


@login_required
def remessa_create(request):
    form = RemessaCreateForm(request.POST or None, initial={"dias_json": "[]"})
    if request.method == "POST" and form.is_valid():
        try:
            dias = _parse_dias(form.cleaned_data["dias_json"])
            textos = {
                "DOCUMENTO_1": form.cleaned_data["texto_documento_1"],
                "DOCUMENTO_2": form.cleaned_data["texto_documento_2"],
                "DOCUMENTO_3": form.cleaned_data["texto_documento_3"],
            }
            remessa = criar_remessa_completa(
                pessoa=form.cleaned_data["pessoa"],
                empresa=form.cleaned_data["empresa"],
                motivo_contratacao=form.cleaned_data["motivo_contratacao"],
                funcao=form.cleaned_data["funcao"],
                data_servico=form.cleaned_data["data_servico"],
                dias_trabalhados=dias,
                textos_por_tipo=textos,
                usuario=request.user,
            )
            messages.success(request, f"Remessa {remessa.numero_formatado} gerada com 3 documentos.")
            return redirect("remessas:detail", pk=remessa.pk)
        except (RemessaError, ValueError, json.JSONDecodeError) as exc:
            messages.error(request, str(exc))
    return render(
        request,
        "remessas/form.html",
        {"form": form, "titulo": "Nova previsão de pagamento"},
    )


@login_required
def remessa_detail(request, pk):
    remessa = get_object_or_404(
        Remessa.objects.select_related("pessoa", "servico__empresa").prefetch_related(
            "documentos", "servico__dias_trabalhados"
        ),
        pk=pk,
    )
    return render(request, "remessas/detail.html", {"remessa": remessa})


@login_required
def remessa_edit(request, pk):
    remessa = get_object_or_404(Remessa.objects.select_related("servico"), pk=pk)
    initial = {
        "pessoa": remessa.pessoa_id,
        "empresa": remessa.servico.empresa_id,
        "data_servico": remessa.servico.data_servico,
        "motivo_contratacao": remessa.servico.motivo_contratacao,
        "funcao": remessa.servico.funcao,
    }
    docs = {d.tipo_documento: d.texto_padrao_id for d in remessa.documentos.all()}
    if docs.get("DOCUMENTO_1"):
        initial["texto_documento_1"] = docs["DOCUMENTO_1"]
    if docs.get("DOCUMENTO_2"):
        initial["texto_documento_2"] = docs["DOCUMENTO_2"]
    if docs.get("DOCUMENTO_3"):
        initial["texto_documento_3"] = docs["DOCUMENTO_3"]

    dias_inicial = json.dumps(
        [
            {
                "data": d.data.isoformat(),
                "quantidade_horas": float(d.quantidade_horas),
                "observacao": d.observacao,
            }
            for d in remessa.servico.dias_trabalhados.all()
        ]
    )
    initial["dias_json"] = dias_inicial
    form = RemessaCreateForm(request.POST or None, initial=initial)
    form.fields["pessoa"].disabled = True
    if request.method == "POST" and form.is_valid():
        try:
            dias = _parse_dias(form.cleaned_data["dias_json"])
            textos = {
                "DOCUMENTO_1": form.cleaned_data["texto_documento_1"],
                "DOCUMENTO_2": form.cleaned_data["texto_documento_2"],
                "DOCUMENTO_3": form.cleaned_data["texto_documento_3"],
            }
            editar_remessa(
                remessa=remessa,
                motivo_contratacao=form.cleaned_data["motivo_contratacao"],
                funcao=form.cleaned_data["funcao"],
                data_servico=form.cleaned_data["data_servico"],
                dias_trabalhados=dias,
                textos_por_tipo=textos,
                usuario=request.user,
                motivo_alteracao=form.cleaned_data.get("motivo_alteracao") or "Edição da remessa",
            )
            messages.success(request, "Remessa atualizada e PDFs regenerados.")
            return redirect("remessas:detail", pk=remessa.pk)
        except (RemessaError, ValueError, json.JSONDecodeError) as exc:
            messages.error(request, str(exc))
    return render(
        request,
        "remessas/form.html",
        {
            "form": form,
            "titulo": f"Editar remessa {remessa.numero_formatado}",
            "remessa": remessa,
        },
    )


@login_required
def remessa_cancel(request, pk):
    remessa = get_object_or_404(Remessa, pk=pk)
    if request.method == "POST":
        remessa.status = Remessa.Status.CANCELADA
        remessa.save(update_fields=["status", "updated_at"])
        messages.success(request, "Remessa cancelada. Numeração preservada.")
        return redirect("remessas:detail", pk=pk)
    return render(request, "remessas/confirm_cancel.html", {"remessa": remessa})
