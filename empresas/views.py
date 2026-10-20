from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.shortcuts import get_object_or_404, redirect, render

from empresas.forms import EmpresaForm
from empresas.models import Empresa


@login_required
def empresa_list(request):
    q = request.GET.get("q", "").strip()
    qs = Empresa.objects.all()
    if q:
        qs = qs.filter(Q(nome__icontains=q) | Q(cnpj__icontains=q.replace(".", "").replace("/", "").replace("-", "")))
    return render(request, "empresas/list.html", {"empresas": qs, "q": q})


@login_required
def empresa_create(request):
    form = EmpresaForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Empresa cadastrada.")
        return redirect("empresas:list")
    return render(request, "empresas/form.html", {"form": form, "titulo": "Nova empresa"})


@login_required
def empresa_edit(request, pk):
    empresa = get_object_or_404(Empresa, pk=pk)
    form = EmpresaForm(request.POST or None, instance=empresa)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Empresa atualizada.")
        return redirect("empresas:list")
    return render(request, "empresas/form.html", {"form": form, "titulo": "Editar empresa", "empresa": empresa})


@login_required
def empresa_detail(request, pk):
    empresa = get_object_or_404(Empresa, pk=pk)
    return render(request, "empresas/detail.html", {"empresa": empresa})


@login_required
def empresa_delete(request, pk):
    empresa = get_object_or_404(Empresa, pk=pk)
    if request.method == "POST":
        if not empresa.pode_excluir():
            messages.error(request, "Empresa vinculada a serviços. Desative em vez de excluir.")
            return redirect("empresas:detail", pk=pk)
        empresa.delete()
        messages.success(request, "Empresa removida.")
        return redirect("empresas:list")
    return render(request, "empresas/confirm_delete.html", {"empresa": empresa})
