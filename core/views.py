from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import redirect, render
from django.utils import timezone

from core.forms import NumeracaoForm
from core.models import SequenciaAnual
from core.services.numeracao import definir_ultimo_numero, obter_ultimo_numero
from documentos.models import Documento
from empresas.models import Empresa
from pessoas.models import Pessoa
from servicos.models import Servico


@login_required
def dashboard(request):
    context = {
        "total_pessoas": Pessoa.objects.filter(ativo=True).count(),
        "total_empresas": Empresa.objects.filter(ativo=True).count(),
        "total_servicos": Servico.objects.count(),
        "total_documentos": Documento.objects.count(),
        "ultimos_servicos": Servico.objects.select_related("pessoa", "empresa")[:8],
    }
    return render(request, "core/dashboard.html", context)


@login_required
def numeracao(request):
    ano = timezone.localdate().year
    inicial = {
        "ultimo_ci": obter_ultimo_numero(SequenciaAnual.Tipo.SERVICO, ano),
        "ultimo_rpa": obter_ultimo_numero(SequenciaAnual.Tipo.RPA, ano),
    }
    if request.method == "POST":
        form = NumeracaoForm(request.POST, ano=ano)
        if form.is_valid():
            definir_ultimo_numero(
                SequenciaAnual.Tipo.SERVICO, form.cleaned_data["ultimo_ci"], ano
            )
            definir_ultimo_numero(
                SequenciaAnual.Tipo.RPA, form.cleaned_data["ultimo_rpa"], ano
            )
            messages.success(
                request,
                "Numeração atualizada. O próximo serviço soma 1 em cada número.",
            )
            return redirect("core:numeracao")
    else:
        form = NumeracaoForm(initial=inicial, ano=ano)

    ultimo_ci = form["ultimo_ci"].value()
    ultimo_rpa = form["ultimo_rpa"].value()
    try:
        proximo_ci = int(ultimo_ci) + 1
        proximo_rpa = int(ultimo_rpa) + 1
    except (TypeError, ValueError):
        proximo_ci = None
        proximo_rpa = None

    return render(
        request,
        "core/numeracao.html",
        {
            "form": form,
            "ano": ano,
            "proximo_ci": f"{proximo_ci:03d}/{ano}" if proximo_ci is not None else "",
            "proximo_rpa": f"{proximo_rpa:03d}/{ano}" if proximo_rpa is not None else "",
        },
    )
