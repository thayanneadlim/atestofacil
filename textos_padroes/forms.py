from django import forms
from textos_padroes.models import TextoPadrao


class TextoPadraoForm(forms.ModelForm):
    class Meta:
        model = TextoPadrao
        fields = ["nome", "descricao", "tipo_documento", "conteudo", "ativo"]
        widgets = {
            "nome": forms.TextInput(attrs={"class": "form-control"}),
            "descricao": forms.Textarea(attrs={"class": "form-control", "rows": 2}),
            "tipo_documento": forms.Select(attrs={"class": "form-select"}),
            "conteudo": forms.Textarea(attrs={"class": "form-control", "rows": 16}),
            "ativo": forms.CheckboxInput(attrs={"class": "form-check-input"}),
        }
