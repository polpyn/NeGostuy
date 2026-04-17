from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('api', '0001_initial'),
    ]

    operations = [
        migrations.AddField(
            model_name='document',
            name='task_id',
            field=models.CharField(blank=True, default='', max_length=255, verbose_name='ID фоновой задачи'),
        ),
    ]
