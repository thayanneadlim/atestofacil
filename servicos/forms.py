from django import forms
from django.utils import timezone

from empresas.models import Empresa
from pessoas.models import Pessoa
from servicos.models import DiaTrabalhado, Servico
from textos_padroes.models import TextoPadrao


class AtestoAssinadoForm(forms.ModelForm):
    class Meta:
        model = Servico
        fields = ["atesto_assinado"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["atesto_assinado"].required = True


class ServicoForm(forms.ModelForm):
    texto_documento_1 = forms.ModelChoiceField(
        queryset=TextoPadrao.objects.none(),
        label="Texto Documento 1",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    texto_documento_2 = forms.ModelChoiceField(
        queryset=TextoPadrao.objects.none(),
        label="Texto Documento 2 (opcional)",
        required=False,
        empty_label="— Não gerar —",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    texto_documento_3 = forms.ModelChoiceField(
        queryset=TextoPadrao.objects.none(),
        label="Texto Documento 3 (opcional)",
        required=False,
        empty_label="— Não gerar —",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    dias_json = forms.CharField(widget=forms.HiddenInput(), initial="[]")
    motivo_alteracao = forms.CharField(
        required=False,
        label="Motivo da alteração",
        widget=forms.TextInput(attrs={"class": "form-control"}),
    )

    class Meta:
        model = Servico
        fields = [
            "empresa",
            "pessoa",
            "data_servico",
            "motivo_contratacao",
            "funcao",
            "assinatura",
        ]
        widgets = {
            "empresa": forms.Select(
                attrs={"class": "form-select", "data-busca-cadastro": "empresa"}
            ),
            "pessoa": forms.Select(
                attrs={"class": "form-select", "data-busca-cadastro": "pessoa"}
            ),
            "data_servico": forms.DateInput(
                format="%Y-%m-%d",
                attrs={"class": "form-control", "type": "date"},
            ),
            "motivo_contratacao": forms.Textarea(attrs={"class": "form-control", "rows": 3}),
            "funcao": forms.TextInput(attrs={"class": "form-control"}),
            "assinatura": forms.Select(attrs={"class": "form-select"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["empresa"].queryset = Empresa.objects.filter(ativo=True)
        self.fields["pessoa"].queryset = Pessoa.objects.filter(ativo=True)
        self.fields["texto_documento_1"].queryset = TextoPadrao.objects.filter(
            ativo=True, tipo_documento="DOCUMENTO_1"
        )
        self.fields["texto_documento_2"].queryset = TextoPadrao.objects.filter(
            ativo=True, tipo_documento="DOCUMENTO_2"
        )
        self.fields["texto_documento_3"].queryset = TextoPadrao.objects.filter(
            ativo=True, tipo_documento="DOCUMENTO_3"
        )
        data_servico = self.fields["data_servico"]
        data_servico.widget.format = "%Y-%m-%d"
        data_servico.input_formats = ["%Y-%m-%d"]
        data_servico.localize = False
        if not self.is_bound:
            self.initial["data_servico"] = timezone.localdate()


class DiaTrabalhadoForm(forms.ModelForm):
    class Meta:
        model = DiaTrabalhado
        fields = ["data", "hora_inicio", "hora_fim", "observacao"]
        widgets = {
            "data": forms.DateInput(attrs={"class": "form-control", "type": "date"}),
            "hora_inicio": forms.TimeInput(attrs={"class": "form-control", "type": "time"}),
            "hora_fim": forms.TimeInput(attrs={"class": "form-control", "type": "time"}),
            "observacao": forms.TextInput(attrs={"class": "form-control"}),
        }
