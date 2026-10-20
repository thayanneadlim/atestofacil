from django.db import migrations, models


def copiar_numeros_rpa(apps, schema_editor):
    Servico = apps.get_model("servicos", "Servico")
    SequenciaAnual = apps.get_model("core", "SequenciaAnual")

    for servico in Servico.objects.all():
        servico.numero_rpa = servico.numero
        servico.ano_rpa = servico.ano
        servico.numero_rpa_documento = servico.numero_documento
        servico.save(update_fields=["numero_rpa", "ano_rpa", "numero_rpa_documento"])

    anos = set(Servico.objects.values_list("ano", flat=True))
    for sequencia in SequenciaAnual.objects.filter(tipo="SERVICO"):
        anos.add(sequencia.ano)
        rpa, criada = SequenciaAnual.objects.get_or_create(
            tipo="RPA",
            ano=sequencia.ano,
            defaults={"ultimo_numero": sequencia.ultimo_numero},
        )
        if not criada and rpa.ultimo_numero < sequencia.ultimo_numero:
            rpa.ultimo_numero = sequencia.ultimo_numero
            rpa.save(update_fields=["ultimo_numero"])

    for ano in anos:
        maior = (
            Servico.objects.filter(ano=ano)
            .order_by("-numero")
            .values_list("numero", flat=True)
            .first()
            or 0
        )
        rpa, criada = SequenciaAnual.objects.get_or_create(
            tipo="RPA",
            ano=ano,
            defaults={"ultimo_numero": maior},
        )
        if not criada and rpa.ultimo_numero < maior:
            rpa.ultimo_numero = maior
            rpa.save(update_fields=["ultimo_numero"])


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0001_initial"),
        ("servicos", "0002_dia_hora_inicio_fim"),
    ]

    operations = [
        migrations.AddField(
            model_name="servico",
            name="ano_rpa",
            field=models.PositiveIntegerField(null=True),
        ),
        migrations.AddField(
            model_name="servico",
            name="numero_rpa",
            field=models.PositiveIntegerField(null=True),
        ),
        migrations.AddField(
            model_name="servico",
            name="numero_rpa_documento",
            field=models.CharField(max_length=20, null=True),
        ),
        migrations.RunPython(copiar_numeros_rpa, migrations.RunPython.noop),
        migrations.AlterField(
            model_name="servico",
            name="ano_rpa",
            field=models.PositiveIntegerField(),
        ),
        migrations.AlterField(
            model_name="servico",
            name="numero_rpa",
            field=models.PositiveIntegerField(),
        ),
        migrations.AlterField(
            model_name="servico",
            name="numero_rpa_documento",
            field=models.CharField(max_length=20, unique=True),
        ),
        migrations.AddConstraint(
            model_name="servico",
            constraint=models.UniqueConstraint(
                fields=("numero_rpa", "ano_rpa"),
                name="uniq_servico_numero_rpa_ano",
            ),
        ),
    ]
