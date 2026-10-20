from django import forms
from django.db.models import Max

from core.models import SequenciaAnual
from servicos.models import Servico


class NumeracaoForm(forms.Form):
    ultimo_ci = forms.IntegerField(
        min_value=0,
        label="Último número de CI e previsão de pagamento",
        widget=forms.NumberInput(attrs={"class": "form-control", "min": "0"}),
    )
    ultimo_rpa = forms.IntegerField(
        min_value=0,
        label="Último número de RPA",
        widget=forms.NumberInput(attrs={"class": "form-control", "min": "0"}),
    )

    def __init__(self, *args, ano: int, **kwargs):
        self.ano = ano
        super().__init__(*args, **kwargs)

    def clean_ultimo_ci(self):
        return self._validar_ultimo(
            self.cleaned_data["ultimo_ci"],
            SequenciaAnual.Tipo.SERVICO,
            "CI e previsão de pagamento",
        )

    def clean_ultimo_rpa(self):
        return self._validar_ultimo(
            self.cleaned_data["ultimo_rpa"],
            SequenciaAnual.Tipo.RPA,
            "RPA",
        )

    def _validar_ultimo(self, valor: int, tipo: str, nome: str) -> int:
        if tipo == SequenciaAnual.Tipo.RPA:
            usado = (
                Servico.objects.filter(ano_rpa=self.ano).aggregate(m=Max("numero_rpa"))["m"]
                or 0
            )
        else:
            usado = (
                Servico.objects.filter(ano=self.ano).aggregate(m=Max("numero"))["m"] or 0
            )
        if valor < usado:
            raise forms.ValidationError(
                f"Já existe {nome} até {usado:03d}. "
                "Informe esse número ou um maior."
            )
        return valor
