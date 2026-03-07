"""
Модели базы данных проекта НеГостуй

Сущности:
- Document — загруженный документ
- ProcessingResult — результат обработки
- GOSTTemplate — шаблон правил ГОСТ
- ErrorStatistic — статистика ошибок
"""

from django.db import models
from django.contrib.auth.models import User


class GOSTTemplate(models.Model):
    """Шаблон правил ГОСТ для проверки"""

    name = models.CharField(
        max_length=200,
        verbose_name="Название шаблона"
    )
    description = models.TextField(
        blank=True,
        verbose_name="Описание"
    )
    rules_json = models.JSONField(
        default=dict,
        verbose_name="Правила в формате JSON",
        help_text="Параметры ГОСТ: шрифт, размер, интервалы, отступы"
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name="Активен"
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Шаблон ГОСТ"
        verbose_name_plural = "Шаблоны ГОСТ"

    def __str__(self):
        return self.name


class Document(models.Model):
    """Загруженный пользователем документ"""

    STATUS_CHOICES = [
        ('uploaded', 'Загружен'),
        ('queued', 'В очереди'),
        ('parsing', 'Парсинг XML'),
        ('classifying', 'Классификация'),
        ('formatting', 'Форматирование'),
        ('completed', 'Завершено'),
        ('error', 'Ошибка'),
    ]

    WORK_TYPE_CHOICES = [
        ('coursework', 'Курсовая работа'),
        ('diploma', 'Дипломная работа'),
        ('lab', 'Лабораторная работа'),
        ('report', 'Отчёт по практике'),
        ('essay', 'Реферат'),
    ]

    # Связь с пользователем (необязательная — можно без регистрации)
    user = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        verbose_name="Пользователь"
    )

    filename = models.CharField(
        max_length=500,
        verbose_name="Имя файла"
    )
    original_file = models.FileField(
        upload_to='uploads/%Y/%m/%d/',
        verbose_name="Исходный файл"
    )
    work_type = models.CharField(
        max_length=20,
        choices=WORK_TYPE_CHOICES,
        default='coursework',
        verbose_name="Тип работы"
    )
    file_size = models.IntegerField(
        default=0,
        verbose_name="Размер файла (байт)"
    )
    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default='uploaded',
        verbose_name="Статус обработки"
    )
    progress = models.IntegerField(
        default=0,
        verbose_name="Прогресс (%)"
    )

    uploaded_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name="Дата загрузки"
    )

    class Meta:
        verbose_name = "Документ"
        verbose_name_plural = "Документы"
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"{self.filename} ({self.get_status_display()})"


class ProcessingResult(models.Model):
    """Результат обработки документа"""

    document = models.OneToOneField(
        Document,
        on_delete=models.CASCADE,
        related_name='result',
        verbose_name="Документ"
    )
    output_file = models.FileField(
        upload_to='results/%Y/%m/%d/',
        null=True,
        blank=True,
        verbose_name="Исправленный файл"
    )

    # Отчёт об ошибках — JSON
    report_json = models.JSONField(
        default=dict,
        verbose_name="Отчёт (JSON)",
        help_text="Список ошибок: [{type, element, was, became}]"
    )

    # Статистика
    total_elements = models.IntegerField(default=0, verbose_name="Всего элементов")
    errors_count = models.IntegerField(default=0, verbose_name="Ошибок")
    warnings_count = models.IntegerField(default=0, verbose_name="Предупреждений")
    headings_count = models.IntegerField(default=0, verbose_name="Заголовков")
    paragraphs_count = models.IntegerField(default=0, verbose_name="Параграфов")
    lists_count = models.IntegerField(default=0, verbose_name="Списков")
    images_count = models.IntegerField(default=0, verbose_name="Изображений")

    grade = models.CharField(
        max_length=50,
        default='',
        verbose_name="Оценка"
    )

    processing_time = models.FloatField(
        default=0,
        verbose_name="Время обработки (сек)"
    )
    processed_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name="Дата обработки"
    )

    class Meta:
        verbose_name = "Результат обработки"
        verbose_name_plural = "Результаты обработки"

    def __str__(self):
        return f"Результат: {self.document.filename} — {self.errors_count} ошибок"


class ErrorStatistic(models.Model):
    """Статистика типичных ошибок (агрегированные данные)"""

    error_type = models.CharField(
        max_length=100,
        verbose_name="Тип ошибки"
    )
    description = models.TextField(
        verbose_name="Описание ошибки"
    )
    count = models.IntegerField(
        default=0,
        verbose_name="Количество"
    )
    work_type = models.CharField(
        max_length=20,
        blank=True,
        verbose_name="Тип работы"
    )
    last_seen = models.DateTimeField(
        auto_now=True,
        verbose_name="Последнее появление"
    )

    class Meta:
        verbose_name = "Статистика ошибок"
        verbose_name_plural = "Статистика ошибок"

    def __str__(self):
        return f"{self.error_type}: {self.count}"