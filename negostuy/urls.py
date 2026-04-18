import re

from django.contrib import admin
from django.urls import path, include, re_path
from django.conf import settings
from django.conf.urls.static import static
from django.views.generic import TemplateView
from django.views.static import serve

urlpatterns = [
    path('', TemplateView.as_view(template_name='index.html'), name='home'),  # ← Главная
    path('admin/', admin.site.urls),
    path('api/', include('api.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

# При DEBUG=False django.conf.urls.static.static() ничего не добавляет (no-op в исходнике Django).
# runserver тогда не отдаёт /static/ — явно вешаем serve.
if not settings.DEBUG and settings.STATICFILES_DIRS:
    _static_prefix = settings.STATIC_URL.lstrip('/')
    urlpatterns += [
        re_path(
            r'^%s(?P<path>.*)$' % re.escape(_static_prefix),
            serve,
            kwargs={'document_root': settings.STATICFILES_DIRS[0]},
        ),
    ]