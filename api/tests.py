from django.contrib import admin
from .models import Document, ProcessingResult, GOSTTemplate, ErrorStatistic


@admin.register(Document)
class DocumentAdmin(admin.ModelAdmin):
    list_display = ['filename', 'status', 'work_type', 'progress', 'uploaded_at']
    list_filter = ['status', 'work_type']
    search_fields = ['filename']


@admin.register(ProcessingResult)
class ProcessingResultAdmin(admin.ModelAdmin):
    list_display = ['document', 'errors_count', 'warnings_count', 'grade', 'processing_time']
    list_filter = ['grade']


@admin.register(GOSTTemplate)
class GOSTTemplateAdmin(admin.ModelAdmin):
    list_display = ['name', 'is_active', 'created_at']


@admin.register(ErrorStatistic)
class ErrorStatisticAdmin(admin.ModelAdmin):
    list_display = ['error_type', 'count', 'work_type', 'last_seen']
    list_filter = ['error_type', 'work_type']