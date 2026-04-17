import os
import time

from celery import shared_task
from django.core.files import File

from core.analyzer import analyze_document
from .models import Document, ProcessingResult


@shared_task(bind=True)
def process_document_task(self, doc_id: int, template_path: str | None, template_is_temp: bool, zachet_number: str) -> dict:
    """
    Фоновая обработка документа:
    - запускает analyze_document
    - сохраняет ProcessingResult
    - обновляет статус Document
    """
    doc = Document.objects.get(id=doc_id)
    doc.status = 'parsing'
    doc.progress = 15
    doc.save(update_fields=['status', 'progress'])

    start_time = time.time()
    try:
        result_data = analyze_document(
            doc.original_file.path,
            template_path=template_path,
            zachet_number=(zachet_number or '').strip() or None,
        )
        processing_time = time.time() - start_time

        proc_result, _ = ProcessingResult.objects.update_or_create(
            document=doc,
            defaults={
                'report_json': result_data.get('report', {}),
                'total_elements': result_data.get('total_elements', 0),
                'errors_count': result_data.get('errors_count', 0),
                'warnings_count': result_data.get('warnings_count', 0),
                'headings_count': result_data.get('headings_count', 0),
                'paragraphs_count': result_data.get('paragraphs_count', 0),
                'lists_count': result_data.get('lists_count', 0),
                'images_count': result_data.get('images_count', 0),
                'grade': result_data.get('grade', ''),
                'processing_time': processing_time,
            },
        )

        output_path = result_data.get('output_path')
        if output_path and os.path.exists(output_path):
            with open(output_path, 'rb') as f:
                proc_result.output_file.save(f"GOST_{doc.filename}", File(f), save=True)

        doc.status = 'completed'
        doc.progress = 100
        doc.save(update_fields=['status', 'progress'])
        return {'ok': True, 'doc_id': doc_id}

    except Exception as exc:
        doc.status = 'error'
        doc.progress = 100
        doc.save(update_fields=['status', 'progress'])
        raise exc
    finally:
        if template_is_temp and template_path and os.path.exists(template_path):
            try:
                os.unlink(template_path)
            except OSError:
                pass
