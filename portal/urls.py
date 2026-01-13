from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    path('login/', views.login_view, name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('api/register/', views.register_view, name='api_register'),
    path('', views.index_view, name='index'),
    path('portal_de_conexoes/', views.portal_de_conexoes_view, name='portal_de_conexoes'),

    path('aprovar_registro/<int:id>/', views.aprovar_registro, name='aprovar_registro'),
    path('rejeitar_registro/<int:id>/', views.rejeitar_registro, name='rejeitar_registro'),
    path('remover_usuario/<int:id>/', views.remover_usuario, name='remover_usuario'),
    path('editar_usuario/<int:id>/', views.editar_usuario, name='editar_usuario'),
    path('manage_service/', views.manage_service, name='manage_service'),
    path('manage_notification/', views.manage_notification, name='manage_notification'),
    path('api/latest-notification/', views.get_latest_notification, name='get_latest_notification'),
    path('api/get_user_profile/', views.get_user_profile, name='get_user_profile'),
    path('api/update_user_profile/', views.update_user_profile, name='update_user_profile'),
    path('api/compromissos/', views.compromisso_list_create_view, name='api_compromissos_list_create'),
    path('api/compromissos/<int:compromisso_id>/', views.compromisso_detail_view, name='api_compromisso_detail'),
    path('api/compromissos/<int:compromisso_id>/share/', views.share_compromisso_view, name='api_compromisso_share'),
    path('api/users/', views.get_users_list, name='api_users_list'),
    path('api/check_session/', views.check_session, name='api_check_session'),
    path('api/hannah/chat/', views.hannah_chat_view, name='hannah_chat'),
]