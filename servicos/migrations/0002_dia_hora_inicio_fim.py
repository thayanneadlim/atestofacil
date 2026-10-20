import datetime

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("servicos", "0001_initial"),
    ]

    operations = [
        migrations.AddField(
            model_name="diatrabalhado",
            name="hora_inicio",
            field=models.TimeField(default=datetime.time(8, 0)),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name="diatrabalhado",
            name="hora_fim",
            field=models.TimeField(default=datetime.time(17, 0)),
            preserve_default=False,
        ),
    ]
