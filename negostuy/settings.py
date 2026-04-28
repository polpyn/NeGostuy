"""
Django settings для проекта НеГостуй
"""

from pathlib import Path
import os
import mimetypes

from dotenv import load_dotenv

# Эти строки помогут туннелю передавать файлы без обрывов
SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')
CSRF_TRUSTED_ORIGINS = [
    'https://*.pinggy.link',
    'https://*.lhr.life',
    'https://*.localhost.run',
]
mimetypes.add_type("text/css", ".css", True)
mimetypes.add_type("text/javascript", ".js", True)

BASE_DIR = Path(__file__).resolve().parent.parent

# Важно: .env должен подхватываться не только через manage.py, но и при запуске
# через Celery worker / gunicorn / wsgi/asgi.
load_dotenv(BASE_DIR / ".env")

SECRET_KEY = os.getenv('SECRET_KEY', 'negostuy-secret-key-change-in-production')

DEBUG = os.getenv('DEBUG', 'True').lower() in ('1', 'true', 'yes', 'on')

# Пустой ALLOWED_HOSTS в .env даёт [] → DisallowedHost / 400. Префикс «.lhr.life» разрешает
# все поддомены туннеля (см. Django validate_host / is_same_domain).
_allowed_raw = (os.getenv('ALLOWED_HOSTS', '*').strip() or '*')
ALLOWED_HOSTS = [h.strip() for h in _allowed_raw.split(',') if h.strip()]
if not ALLOWED_HOSTS:
    ALLOWED_HOSTS = ['*']
elif ALLOWED_HOSTS != ['*']:
    # С точкой в начале — все поддомены (Django 5.1+: без точки только exact match).
    for _tunnel in ('.lhr.life', '.localhost.run'):
        if _tunnel not in ALLOWED_HOSTS:
            ALLOWED_HOSTS.append(_tunnel)

# ============================================================
# ПРИЛОЖЕНИЯ
# ============================================================

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',

    # Сторонние
    'rest_framework',
    'corsheaders',

    # Наше приложение (ApiConfig — PRAGMA WAL для SQLite при нескольких клиентах)
    'api.apps.ApiConfig',
]

MIDDLEWARE = [
    'negostuy.middleware.RequestLogMiddleware',
    'corsheaders.middleware.CorsMiddleware',      # CORS — чтобы frontend мог обращаться
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
]

ROOT_URLCONF = 'negostuy.urls'

STATIC_URL = '/static/'
STATICFILES_DIRS = [os.path.join(BASE_DIR, 'static')]

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [os.path.join(BASE_DIR, 'templates')],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

# ============================================================
# БАЗА ДАННЫХ
# ============================================================
# Локально — SQLite по умолчанию. Для VPS/compose используйте PostgreSQL через env.
#
# В .env из Docker Compose часто DB_HOST=postgres — это имя сервиса, оно не резолвится
# на голой машине (runserver без контейнера). Тогда переключаемся на SQLite.
# В контейнере есть /.dockerenv; при нестандартном окружении задайте DB_USE_COMPOSE_NETWORK=1.
_POSTGRES = 'django.db.backends.postgresql'
_db_engine = os.getenv('DB_ENGINE', 'django.db.backends.sqlite3')
_db_host = (os.getenv('DB_HOST', '127.0.0.1') or '').strip().lower()
_in_container = os.path.exists('/.dockerenv')
_compose_network = os.getenv('DB_USE_COMPOSE_NETWORK', '').lower() in ('1', 'true', 'yes', 'on')
if (
    _db_engine == _POSTGRES
    and _db_host == 'postgres'
    and not _in_container
    and not _compose_network
):
    _db_engine = 'django.db.backends.sqlite3'

DB_ENGINE = _db_engine
if DB_ENGINE == 'django.db.backends.sqlite3':
    DATABASES = {
        'default': {
            'ENGINE': DB_ENGINE,
            'NAME': BASE_DIR / 'db.sqlite3',
            # Два клиента по туннелю одновременно шлют /api/upload/ — без WAL/таймаута часто «database is locked» → 500.
            'OPTIONS': {
                'timeout': 30,
            },
        }
    }
else:
    DATABASES = {
        'default': {
            'ENGINE': DB_ENGINE,
            'NAME': os.getenv('DB_NAME', 'negostuy'),
            'USER': os.getenv('DB_USER', 'negostuy'),
            'PASSWORD': os.getenv('DB_PASSWORD', 'negostuy'),
            'HOST': os.getenv('DB_HOST', '127.0.0.1'),
            'PORT': os.getenv('DB_PORT', '5432'),
        }
    }

# PostgreSQL (когда понадобится):
# DATABASES = {
#     'default': {
#         'ENGINE': 'django.db.backends.postgresql',
#         'NAME': 'negostuy_db',
#         'USER': 'postgres',
#         'PASSWORD': 'dasha',
#         'HOST': '127.0.0.1',
#         'PORT': '5432',
#     }
# }

# ============================================================
# CORS — разрешаем запросы с фронтенда
# ============================================================

CORS_ALLOW_ALL_ORIGINS = True  # Для разработки

# ============================================================
# REST FRAMEWORK
# ============================================================

REST_FRAMEWORK = {
    'DEFAULT_AUTHENTICATION_CLASSES': [
        'rest_framework.authentication.SessionAuthentication',
        'rest_framework.authentication.TokenAuthentication',
    ],
    'DEFAULT_PERMISSION_CLASSES': [
        'rest_framework.permissions.AllowAny',  # Пока без обязательной авторизации
    ],
}

# ============================================================
# CELERY / REDIS
# ============================================================
CELERY_BROKER_URL = os.getenv('CELERY_BROKER_URL', 'redis://127.0.0.1:6379/0')
CELERY_RESULT_BACKEND = os.getenv('CELERY_RESULT_BACKEND', CELERY_BROKER_URL)
CELERY_TASK_TRACK_STARTED = True
CELERY_TASK_TIME_LIMIT = int(os.getenv('CELERY_TASK_TIME_LIMIT', '900'))
ENABLE_ASYNC_PROCESSING = os.getenv('ENABLE_ASYNC_PROCESSING', '0').lower() in ('1', 'true', 'yes', 'on')

# ============================================================
# AI POSTPROCESS (Gemini / Ollama)
# ============================================================
AI_POSTPROCESS_ENABLED = os.getenv('AI_POSTPROCESS_ENABLED', '1')

# Провайдер: 'gemini' (по умолчанию, если задан ключ), 'openrouter', 'gigachat', 'ollama' или 'auto'.
AI_PROVIDER = os.getenv('AI_PROVIDER', 'auto').strip().lower()

# AI_FALLBACK_DISABLED=1 — использовать только primary провайдера, не пытаться резервных.
# Полезно, когда primary стабильно отвечает, а резерв (например, Gemini) сейчас лежит
# и тянет ретраями по 60 сек, превращая один запрос в 30 минут ожидания.
AI_FALLBACK_DISABLED = os.getenv('AI_FALLBACK_DISABLED', '0')

# При primary=gigachat по умолчанию НЕ ходим в Gemini как резерв (см. _fallback_chain).
# Если хочется обратного — выставите AI_GEMINI_FALLBACK_FROM_GIGACHAT=1.
AI_GEMINI_FALLBACK_FROM_GIGACHAT = os.getenv('AI_GEMINI_FALLBACK_FROM_GIGACHAT', '0')

# Google Gemini (AI Studio). Бесплатный лимит для gemini-2.5-flash.
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY', '')
GEMINI_MODEL = os.getenv('GEMINI_MODEL', 'gemini-2.5-flash')
GEMINI_TIMEOUT = int(os.getenv('GEMINI_TIMEOUT', '120'))
GEMINI_BASE_URL = os.getenv(
    'GEMINI_BASE_URL', 'https://generativelanguage.googleapis.com/v1beta'
)
# HTTP(S) прокси для Gemini (нужен из РФ). Пример: http://user:pass@host:port
GEMINI_PROXY = os.getenv('GEMINI_PROXY', '').strip()
# Повтор при 503/429 и т.п. (иначе «первое окно» с началом документа может остаться пустым).
# Дефолт 2 — чтобы при недоступном Gemini весь прогон не залипал на 30+ минут
# (8 ретраев × до 60 секунд таймаута на каждое из 4 окон).
GEMINI_MAX_RETRIES = int(os.getenv('GEMINI_MAX_RETRIES', '2'))
GEMINI_RETRY_BACKOFF_SEC = float(os.getenv('GEMINI_RETRY_BACKOFF_SEC', '2'))
GEMINI_RETRY_MAX_WAIT_SEC = float(os.getenv('GEMINI_RETRY_MAX_WAIT_SEC', '20'))

# OpenRouter — работает без VPN из РФ, даёт бесплатный лимит.
# Ключ: https://openrouter.ai/keys
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY', '').strip()
OPENROUTER_MODEL = os.getenv('OPENROUTER_MODEL', 'google/gemini-2.5-flash')
OPENROUTER_TIMEOUT = int(os.getenv('OPENROUTER_TIMEOUT', '120'))
OPENROUTER_BASE_URL = os.getenv(
    'OPENROUTER_BASE_URL', 'https://openrouter.ai/api/v1'
)

# GigaChat (Сбер)
GIGACHAT_AUTH_KEY = os.getenv('GIGACHAT_AUTH_KEY', '').strip()
GIGACHAT_SCOPE = os.getenv('GIGACHAT_SCOPE', 'GIGACHAT_API_PERS').strip()
GIGACHAT_MODEL = os.getenv('GIGACHAT_MODEL', 'GigaChat').strip()
# Таймаут на один HTTP-запрос к GigaChat. 60 сек достаточно для окна 50 абзацев;
# если зависает — лучше быстро отвалиться, чем держать пользователя 2 минуты.
GIGACHAT_TIMEOUT = int(os.getenv('GIGACHAT_TIMEOUT', '60'))
GIGACHAT_BASE_URL = os.getenv(
    'GIGACHAT_BASE_URL', 'https://gigachat.devices.sberbank.ru/api/v1'
).rstrip('/')
GIGACHAT_OAUTH_URL = os.getenv(
    'GIGACHAT_OAUTH_URL', 'https://ngw.devices.sberbank.ru:9443/api/v2/oauth'
).rstrip('/')
GIGACHAT_INSECURE_SSL = os.getenv('GIGACHAT_INSECURE_SSL', '0').lower() in ('1', 'true', 'yes', 'on')

# Локальный Ollama — используется как fallback.
OLLAMA_BASE_URL = os.getenv('OLLAMA_BASE_URL', 'http://127.0.0.1:11434')
OLLAMA_MODEL = os.getenv('OLLAMA_MODEL', 'qwen2.5-coder:14b')
OLLAMA_TIMEOUT = int(os.getenv('OLLAMA_TIMEOUT', '240'))

# Размер окна кандидатов (в абзацах), которое отправляется в LLM за раз.
AI_POSTPROCESS_WINDOW_SIZE = int(os.getenv('AI_POSTPROCESS_WINDOW_SIZE', '50'))
# Перекрытие окон, чтобы модель видела контекст предыдущего окна.
AI_POSTPROCESS_WINDOW_OVERLAP = int(os.getenv('AI_POSTPROCESS_WINDOW_OVERLAP', '5'))
# Максимальное число окон за один прогон (защита от бесконечных документов).
# Шаг смещения окна = WINDOW_SIZE - OVERLAP (напр. 45). При MAX_WINDOWS=40 покрыто ~45*39+50 абзацев.
AI_POSTPROCESS_MAX_WINDOWS = int(os.getenv('AI_POSTPROCESS_MAX_WINDOWS', '40'))
# Пауза между окнами (сек), снижает 429 при бесплатном лимите RPM. 0 = без паузы.
AI_POSTPROCESS_INTER_WINDOW_SEC = float(os.getenv('AI_POSTPROCESS_INTER_WINDOW_SEC', '0'))
# 1 = сначала только окна с нумерацией/маркерами (дешевле; ввод без списков может не анализироваться).
# 0 = все окна подряд (рекомендуется для конспектов).
AI_POSTPROCESS_LIST_ONLY_WINDOWS = os.getenv(
    'AI_POSTPROCESS_LIST_ONLY_WINDOWS', '0'
).lower() in ('1', 'true', 'yes', 'on')
# Legacy-параметр, если где-то ещё читается.
AI_POSTPROCESS_MAX_CANDIDATES = int(os.getenv('AI_POSTPROCESS_MAX_CANDIDATES', '30'))

AI_POSTPROCESS_DEBUG = os.getenv('AI_POSTPROCESS_DEBUG', '1')
AI_POSTPROCESS_DEBUG_DIR = os.getenv('AI_POSTPROCESS_DEBUG_DIR', '')
# Подробная трассировка: полный промпт/ответ в JSON и опционально файлы wN_prompt.txt / wN_response.txt
AI_POSTPROCESS_TRACE = os.getenv('AI_POSTPROCESS_TRACE', '0')
AI_POSTPROCESS_TRACE_FILES = os.getenv('AI_POSTPROCESS_TRACE_FILES', '0')
AI_POSTPROCESS_TRACE_PROMPT_HEAD = int(os.getenv('AI_POSTPROCESS_TRACE_PROMPT_HEAD', '8000'))
AI_POSTPROCESS_TRACE_PROMPT_TAIL = int(os.getenv('AI_POSTPROCESS_TRACE_PROMPT_TAIL', '4000'))
AI_POSTPROCESS_TRACE_RESPONSE_MAX = int(os.getenv('AI_POSTPROCESS_TRACE_RESPONSE_MAX', '200000'))
AI_POSTPROCESS_TRACE_CONSOLE = os.getenv('AI_POSTPROCESS_TRACE_CONSOLE', '1')
AI_POSTPROCESS_HEURISTIC_FALLBACK = os.getenv('AI_POSTPROCESS_HEURISTIC_FALLBACK', '1')
AI_POSTPROCESS_MIN_CONFIDENCE = float(os.getenv('AI_POSTPROCESS_MIN_CONFIDENCE', '0.35'))

# ============================================================
# ФАЙЛЫ
# ============================================================

MEDIA_ROOT = os.path.join(BASE_DIR, 'media')
MEDIA_URL = '/media/'

# Максимальный размер загрузки — 16 МБ
DATA_UPLOAD_MAX_MEMORY_SIZE = 16 * 1024 * 1024
FILE_UPLOAD_MAX_MEMORY_SIZE = 16 * 1024 * 1024

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

LANGUAGE_CODE = 'ru-ru'
TIME_ZONE = 'Europe/Moscow'