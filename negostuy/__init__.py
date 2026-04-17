try:
    from .celery import app as celery_app
except Exception:  # Celery может быть не установлен в локальном окружении
    celery_app = None

__all__ = ('celery_app',)
