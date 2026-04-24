from django.apps import AppConfig
from django.db.backends.signals import connection_created


def _sqlite_connection_init(sender, connection, **kwargs):
    if connection.vendor != 'sqlite':
        return
    with connection.cursor() as cursor:
        cursor.execute('PRAGMA journal_mode=WAL;')
        cursor.execute('PRAGMA synchronous=NORMAL;')
        cursor.execute('PRAGMA busy_timeout=30000;')


class ApiConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'api'

    def ready(self):
        connection_created.connect(
            _sqlite_connection_init, dispatch_uid='api.sqlite_wal_busy'
        )
