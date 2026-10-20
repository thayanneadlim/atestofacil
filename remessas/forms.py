from django import forms
from pessoas.models import Pessoa
from empresas.models import Empresa
from textos_padroes.models import TextoPadrao


class RemessaCreateForm(forms.Form):
    pessoa = forms.ModelChoiceField(
        queryset=Pessoa.objects.filter(ativo=True),
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    empresa = forms.ModelChoiceField(
        queryset=Empresa.objects.filter(ativo=True),
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    data_servico = forms.DateField(
        widget=forms.DateInput(attrs={"class": "form-control", "type": "date"}),
    )
    motivo_contratacao = forms.CharField(
        widget=forms.Textarea(attrs={"class": "form-control", "rows": 3}),
    )
    funcao = forms.CharField(
        max_length=255,
        widget=forms.TextInput(attrs={"class": "form-control"}),
    )
    texto_documento_1 = forms.ModelChoiceField(
        queryset=TextoPadrao.objects.filter(ativo=True, tipo_documento="DOCUMENTO_1"),
        label="Texto Documento 1",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    texto_documento_2 = forms.ModelChoiceField(
        queryset=TextoPadrao.objects.filter(ativo=True, tipo_documento="DOCUMENTO_2"),
        label="Texto Documento 2",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    texto_documento_3 = forms.ModelChoiceField(
        queryset=TextoPadrao.objects.filter(ativo=True, tipo_documento="DOCUMENTO_3"),
        label="Texto Documento 3",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    dias_json = forms.CharField(widget=forms.HiddenInput())
    motivo_alteracao = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        label="Motivo da alteração",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["texto_documento_1"].queryset = TextoPadrao.objects.filter(
            ativo=True, tipo_documento="DOCUMENTO_1"
        )
        self.fields["texto_documento_2"].queryset = TextoPadrao.objects.filter(
            ativo=True, tipo_documento="DOCUMENTO_2"
        )
        self.fields["texto_documento_3"].queryset = TextoPadrao.objects.filter(
            ativo=True, tipo_documento="DOCUMENTO_3"
        )


class RemessaFilterForm(forms.Form):
    q = forms.CharField(
        required=False,
        label="Busca",
        widget=forms.TextInput(attrs={"class": "form-control", "placeholder": "Nome, CPF ou número"}),
    )
    empresa = forms.ModelChoiceField(
        required=False,
        queryset=Empresa.objects.all(),
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    data_inicio = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={"class": "form-control", "type": "date"}),
    )
    data_fim = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={"class": "form-control", "type": "date"}),
    )
