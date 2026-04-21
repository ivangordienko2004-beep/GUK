from django.contrib.auth import views as auth_views
from django.urls import path

from . import views

urlpatterns = [
    path('login/', auth_views.LoginView.as_view(template_name='core/login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
    path(
        'password-change/',
        auth_views.PasswordChangeView.as_view(template_name='core/password_change.html'),
        name='password_change',
    ),
    path(
        'password-change/done/',
        auth_views.PasswordChangeDoneView.as_view(template_name='core/password_change_done.html'),
        name='password_change_done',
    ),
    path('', views.dashboard, name='dashboard'),
    path('questionnaire/', views.questionnaire, name='questionnaire'),
    path('upload/', views.upload_files, name='upload_files'),
    path('download/merged/', views.download_merged, name='download_merged'),
    path('editor/', views.excel_editor, name='excel_editor'),
    path('editor/save/', views.save_excel_editor, name='save_excel_editor'),
    path('decode/', views.decode_vus, name='decode_vus'),
    path('report/', views.create_report, name='create_report'),
]
