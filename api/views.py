"""
API views для проекта НеГостуй

Эндпоинты:
- POST /api/upload/          — загрузка документа
- GET  /api/status/<id>/     — статус обработки
- GET  /api/report/<id>/     — отчёт об ошибках
- GET  /api/download/<id>/   — скачивание результата
- POST /api/auth/register/   — регистрация
- POST /api/auth/login/      — вход
- GET  /api/health/          — проверка сервера
"""

import os
import time
import json

from django.http import FileResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout

from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import AllowAny, IsAuthenticated
from rest_framework.response import Response
from rest_framework import status

from .models import Document, ProcessingResult, GOSTTemplate, ErrorStatistic
from core.analyzer import analyze_document


# ============================================================
# HEALTH CHECK
# ============================================================

@api_view(['GET'])
@permission_classes([AllowAny])
def health_check(request):
    """Проверка работоспособности сервера"""
    return Response({
        'status': 'ok',
        'message': 'НеГостуй API работает',
        'version': '1.0'
    })


# ============================================================
# ЗАГРУЗКА И ОБРАБОТКА ДОКУМЕНТА
# ============================================================

@api_view(['POST'])
@permission_classes([AllowAny])
def upload_document(request):
    """
    POST /api/upload/
    Принимает .docx файл + необязательную рамку (template).
    """

    if 'file' not in request.FILES:
        return Response(
            {'error': 'Файл не найден. Отправьте файл в поле "file"'},
            status=status.HTTP_400_BAD_REQUEST
        )

    file = request.FILES['file']

    # Валидация документа
    if not file.name.endswith('.docx'):
        return Response({'error': 'Разрешены только .docx'}, status=status.HTTP_400_BAD_REQUEST)
    if file.size > 16 * 1024 * 1024:
        return Response({'error': 'Максимум 16 МБ'}, status=status.HTTP_400_BAD_REQUEST)
    if file.size == 0:
        return Response({'error': 'Файл пустой'}, status=status.HTTP_400_BAD_REQUEST)

    # Рамка (необязательно)
    template_file = request.FILES.get('template', None)
    template_path = None

    if template_file:
        if not template_file.name.endswith('.docx'):
            return Response({'error': 'Рамка должна быть .docx'}, status=status.HTTP_400_BAD_REQUEST)

        # Сохраняем рамку во временный файл
        import tempfile
        tmp_template = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        for chunk in template_file.chunks():
            tmp_template.write(chunk)
        tmp_template.close()
        template_path = tmp_template.name
        print(f"🖼️ Рамка сохранена: {template_path}")

    work_type = request.data.get('work_type', 'coursework')

    # Сохраняем документ в БД
    doc = Document.objects.create(
        user=request.user if request.user.is_authenticated else None,
        filename=file.name,
        original_file=file,
        file_size=file.size,
        work_type=work_type,
        status='uploaded'
    )

    try:
        doc.status = 'parsing'
        doc.progress = 10
        doc.save()

        start_time = time.time()

        # Вызываем ядро с путём к рамке
        result_data = analyze_document(
            doc.original_file.path,
            template_path=template_path
        )

        processing_time = time.time() - start_time

        # Сохраняем результат
        proc_result = ProcessingResult.objects.create(
            document=doc,
            report_json=result_data.get('report', {}),
            total_elements=result_data.get('total_elements', 0),
            errors_count=result_data.get('errors_count', 0),
            warnings_count=result_data.get('warnings_count', 0),
            headings_count=result_data.get('headings_count', 0),
            paragraphs_count=result_data.get('paragraphs_count', 0),
            lists_count=result_data.get('lists_count', 0),
            images_count=result_data.get('images_count', 0),
            grade=result_data.get('grade', ''),
            processing_time=processing_time,
        )

        # Сохраняем исправленный файл
        if result_data.get('output_path') and os.path.exists(result_data['output_path']):
            from django.core.files import File
            with open(result_data['output_path'], 'rb') as f:
                proc_result.output_file.save(f"GOST_{doc.filename}", File(f))

        doc.status = 'completed'
        doc.progress = 100
        doc.save()

        return Response({
            'success': True,
            'document_id': doc.id,
            'document_name': doc.filename,
            'status': 'completed',
            'summary': {
                'total_elements': proc_result.total_elements,
                'headings': proc_result.headings_count,
                'lists': proc_result.lists_count,
                'texts': proc_result.paragraphs_count,
                'images': proc_result.images_count,
                'errors_count': proc_result.errors_count,
                'warnings_count': proc_result.warnings_count,
                'perfect_elements': proc_result.total_elements - proc_result.errors_count,
                'elements_with_errors': proc_result.errors_count,
                'grade': proc_result.grade,
                'processing_time': round(processing_time, 2),
                'has_template': template_path is not None,
            },
            'elements': result_data.get('elements_detail', []),
        })

    except Exception as e:
        doc.status = 'error'
        doc.save()
        import traceback
        traceback.print_exc()
        return Response(
            {'success': False, 'error': str(e)},
            status=status.HTTP_500_INTERNAL_SERVER_ERROR
        )

    finally:
        # Удаляем временный файл рамки
        if template_path and os.path.exists(template_path):
            try:
                os.unlink(template_path)
            except:
                pass

# ============================================================
# СТАТУС ОБРАБОТКИ
# ============================================================

@api_view(['GET'])
@permission_classes([AllowAny])
def document_status(request, doc_id):
    """
    GET /api/status/<id>/
    Возвращает текущий статус обработки документа.
    """
    try:
        doc = Document.objects.get(id=doc_id)
        return Response({
            'document_id': doc.id,
            'filename': doc.filename,
            'status': doc.status,
            'progress': doc.progress,
            'status_display': doc.get_status_display(),
        })
    except Document.DoesNotExist:
        return Response(
            {'error': 'Документ не найден'},
            status=status.HTTP_404_NOT_FOUND
        )


# ============================================================
# ОТЧЁТ ОБ ОШИБКАХ
# ============================================================

@api_view(['GET'])
@permission_classes([AllowAny])
def document_report(request, doc_id):
    """
    GET /api/report/<id>/
    Возвращает подробный отчёт об ошибках.
    """
    try:
        doc = Document.objects.get(id=doc_id)

        if doc.status != 'completed':
            return Response({
                'error': 'Документ ещё обрабатывается',
                'status': doc.status,
                'progress': doc.progress,
            })

        result = doc.result

        return Response({
            'document_id': doc.id,
            'filename': doc.filename,
            'report': result.report_json,
            'summary': {
                'total_elements': result.total_elements,
                'errors_count': result.errors_count,
                'warnings_count': result.warnings_count,
                'grade': result.grade,
                'processing_time': result.processing_time,
            }
        })

    except Document.DoesNotExist:
        return Response(
            {'error': 'Документ не найден'},
            status=status.HTTP_404_NOT_FOUND
        )
    except ProcessingResult.DoesNotExist:
        return Response(
            {'error': 'Результат ещё не готов'},
            status=status.HTTP_404_NOT_FOUND
        )


# ============================================================
# СКАЧИВАНИЕ РЕЗУЛЬТАТА
# ============================================================

@api_view(['GET'])
@permission_classes([AllowAny])
def download_result(request, doc_id):
    """
    GET /api/download/<id>/
    Возвращает исправленный .docx файл.
    """
    try:
        doc = Document.objects.get(id=doc_id)

        if doc.status != 'completed':
            return Response(
                {'error': 'Документ ещё обрабатывается'},
                status=status.HTTP_400_BAD_REQUEST
            )

        result = doc.result

        if not result.output_file:
            return Response(
                {'error': 'Исправленный файл не найден'},
                status=status.HTTP_404_NOT_FOUND
            )

        return FileResponse(
            result.output_file.open('rb'),
            as_attachment=True,
            filename=f"GOST_{doc.filename}"
        )

    except (Document.DoesNotExist, ProcessingResult.DoesNotExist):
        return Response(
            {'error': 'Документ не найден'},
            status=status.HTTP_404_NOT_FOUND
        )


# ============================================================
# АУТЕНТИФИКАЦИЯ
# ============================================================

@api_view(['POST'])
@permission_classes([AllowAny])
def register_user(request):
    """
    POST /api/auth/register/
    Регистрация нового пользователя.
    """
    username = request.data.get('username', '').strip()
    email = request.data.get('email', '').strip()
    password = request.data.get('password', '')

    # Валидация
    if not username:
        return Response(
            {'error': 'Укажите имя пользователя'},
            status=status.HTTP_400_BAD_REQUEST
        )
    if not password or len(password) < 6:
        return Response(
            {'error': 'Пароль должен содержать минимум 6 символов'},
            status=status.HTTP_400_BAD_REQUEST
        )
    if User.objects.filter(username=username).exists():
        return Response(
            {'error': 'Пользователь с таким именем уже существует'},
            status=status.HTTP_400_BAD_REQUEST
        )
    if email and User.objects.filter(email=email).exists():
        return Response(
            {'error': 'Email уже используется'},
            status=status.HTTP_400_BAD_REQUEST
        )

    user = User.objects.create_user(
        username=username,
        email=email,
        password=password
    )

    return Response({
        'success': True,
        'message': 'Регистрация успешна',
        'user_id': user.id,
        'username': user.username,
    }, status=status.HTTP_201_CREATED)


@api_view(['POST'])
@permission_classes([AllowAny])
def login_user(request):
    """
    POST /api/auth/login/
    Вход пользователя.
    """
    username = request.data.get('username', '')
    password = request.data.get('password', '')

    user = authenticate(request, username=username, password=password)

    if user is not None:
        login(request, user)
        return Response({
            'success': True,
            'message': 'Вход выполнен',
            'user_id': user.id,
            'username': user.username,
        })
    else:
        return Response(
            {'error': 'Неверный логин или пароль'},
            status=status.HTTP_401_UNAUTHORIZED
        )


@api_view(['POST'])
def logout_user(request):
    """POST /api/auth/logout/"""
    logout(request)
    return Response({'success': True, 'message': 'Выход выполнен'})


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ
# ============================================================

# ============================================================
# ТИТУЛЬНЫЙ ЛИСТ
# ============================================================

@csrf_exempt
def generate_title_page(request):
    """
    POST /api/title-page/
    Генерирует титульный лист и возвращает .docx файл.
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'Только POST'}, status=405)

    template_path = None
    output_path = None

    try:
        data = {
            'institution':    request.POST.get('institution', ''),
            'work_title':     request.POST.get('work_title', 'ОТЧЕТ ПО ПРАКТИЧЕСКОЙ РАБОТЕ'),
            'work_number':    request.POST.get('work_number', ''),
            'specialty_code': request.POST.get('specialty_code', ''),
            'specialty_name': request.POST.get('specialty_name', ''),
            'subject':        request.POST.get('subject', ''),
            'group':          request.POST.get('group', ''),
            'student_id':     request.POST.get('student_id', ''),
            'student_name':   request.POST.get('student_name', ''),
            'teacher_name':   request.POST.get('teacher_name', ''),
            'city':           request.POST.get('city', 'Красноярск'),
            'year':           request.POST.get('year', '2026'),
        }

        # Рамка (необязательно)
        template_file = request.FILES.get('template')
        if template_file:
            if not template_file.name.endswith('.docx'):
                return JsonResponse({'error': 'Рамка должна быть .docx'}, status=400)
            import tempfile as _tf
            tmp = _tf.NamedTemporaryFile(delete=False, suffix='.docx')
            for chunk in template_file.chunks():
                tmp.write(chunk)
            tmp.close()
            template_path = tmp.name

        # Генерация
        import tempfile as _tf
        fd, output_path = _tf.mkstemp(suffix='.docx', prefix='title_')
        os.close(fd)

        from core.title_page import create_title_page
        create_title_page(data, output_path, template_path)

        # Читаем файл и возвращаем
        with open(output_path, 'rb') as f:
            content = f.read()

        from django.http import HttpResponse as HR
        response = HR(
            content,
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        work_num = data.get('work_number', '')
        fname = f'Титульный_лист_N{work_num}.docx' if work_num else 'Титульный_лист.docx'
        response['Content-Disposition'] = f'attachment; filename="{fname}"'
        return response

    except Exception as e:
        import traceback
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)

    finally:
        if template_path and os.path.exists(template_path):
            try:
                os.unlink(template_path)
            except:
                pass
        if output_path and os.path.exists(output_path):
            try:
                os.unlink(output_path)
            except:
                pass

def _update_error_stats(report, work_type):
    """Обновляет агрегированную статистику ошибок"""
    if not isinstance(report, dict):
        return

    errors = report.get('errors', [])
    for err in errors:
        err_type = err.get('type', 'unknown')
        stat, _ = ErrorStatistic.objects.get_or_create(
            error_type=err_type,
            work_type=work_type,
            defaults={'description': err.get('description', '')}
        )
        stat.count += 1
        stat.save()