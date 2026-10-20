from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db.models import Q
from django.shortcuts import get_object_or_404, redirect, render

from pessoas.forms import PessoaForm
from pessoas.models import Pessoa


@login_required
def pessoa_list(request):
    q = request.GET.get("q", "").strip()
    qs = Pessoa.objects.all()
    if q:
        qs = qs.filter(Q(nome__icontains=q) | Q(cpf__icontains=q.replace(".", "").replace("-", "")))
    return render(request, "pessoas/list.html", {"pessoas": qs, "q": q})


@login_required
def pessoa_create(request):
    form = PessoaForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Pessoa cadastrada com sucesso.")
        return redirect("pessoas:list")
    return render(request, "pessoas/form.html", {"form": form, "titulo": "Nova pessoa"})


@login_required
def pessoa_edit(request, pk):
    pessoa = get_object_or_404(Pessoa, pk=pk)
    form = PessoaForm(request.POST or None, instance=pessoa)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Pessoa atualizada.")
        return redirect("pessoas:detail", pk=pessoa.pk)
    return render(request, "pessoas/form.html", {"form": form, "titulo": "Editar pessoa", "pessoa": pessoa})


@login_required
def pessoa_detail(request, pk):
    pessoa = get_object_or_404(Pessoa, pk=pk)
    return render(request, "pessoas/detail.html", {"pessoa": pessoa})
