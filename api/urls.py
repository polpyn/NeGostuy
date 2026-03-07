from django.urls import path
from . import views

urlpatterns = [
    # Здоровье сервера
    path('health/', views.health_check, name='health'),

    # Документы
    path('upload/', views.upload_document, name='upload'),
    path('status/<int:doc_id>/', views.document_status, name='status'),
    path('report/<int:doc_id>/', views.document_report, name='report'),
    path('download/<int:doc_id>/', views.download_result, name='download'),

    # Титульный лист                          ← НОВОЕ
    path('title-page/', views.generate_title_page, name='title_page'),

    # Аутентификация
    path('auth/register/', views.register_user, name='register'),
    path('auth/login/', views.login_user, name='login'),
    path('auth/logout/', views.logout_user, name='logout'),
]