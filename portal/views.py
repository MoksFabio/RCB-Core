from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse, HttpResponseForbidden, HttpResponseBadRequest
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required, user_passes_test
from django.db import transaction
from django.views.decorators.http import require_POST, require_GET
from django.views.decorators.csrf import csrf_exempt

from django.urls import reverse
from .models import User, UserProfile, Compromisso, SystemService, GlobalNotification, NotificationUserStatus
import json
from utils.hannah import get_hannah_response
from datetime import datetime, date
from django.contrib import messages

from . import servicos

def index_view(request):
    return redirect(reverse('login'))

@csrf_exempt
def login_view(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            username = data.get('username')
            password = data.get('password')
        except json.JSONDecodeError:
            return JsonResponse({"status": "error", "message": "Requisição inválida."}, status=400)

        if not username or not password:
            return JsonResponse({"status": "error", "message": "Usuário e senha são obrigatórios."}, status=400)

        user = authenticate(request, username=username, password=password)

        if user is not None:
            if user.status != 'aprovado':
                return JsonResponse({'status': 'error', 'message': 'Seu registro ainda não foi aprovado.'}, status=403)
            
            login(request, user)
            request.session['username'] = user.username
            return JsonResponse({'status': 'success', 'redirect_url': reverse('portal_de_conexoes')})
        else:
            return JsonResponse({'status': 'error', 'message': 'Usuário ou senha inválidos.'}, status=401)

    return render(request, 'login.html')

@csrf_exempt
@require_POST
def register_view(request):
    try:
        data = json.loads(request.body)
        username = data.get('username')
        password = data.get('password')
        confirm_password = data.get('confirm_password')
    except json.JSONDecodeError:
        return JsonResponse({"status": "error", "message": "Requisição inválida."}, status=400)

    if not username or not password or not confirm_password:
        return JsonResponse({"status": "error", "message": "Todos os campos são obrigatórios."}, status=400)


    
    if password != confirm_password:
        return JsonResponse({"status": "error", "message": "As senhas não correspondem."}, status=400)

    try:
        servicos.create_user_service(username, password)
        return JsonResponse({"status": "success", "message": "Registro realizado com sucesso! Aguarde a aprovação."}, status=201)
    
    except ValueError as e:
        return JsonResponse({"status": "error", "message": str(e)}, status=409)
    except Exception as e:
        return JsonResponse({"status": "error", "message": f"Erro interno: {str(e)}"}, status=500)



def check_session(request):
    return JsonResponse({'is_authenticated': request.user.is_authenticated})

from django.views.decorators.cache import never_cache


@never_cache
@login_required
def portal_de_conexoes_view(request):
    context = {
        'username': request.user.username,
        'is_super_admin': request.user.is_superuser
    }
    
    # System Services for Widget
    context['services'] = SystemService.objects.all()
    
    # Notifications for Modal handled by context_processor
    # context['notifications'] = GlobalNotification.objects.filter(is_active=True).all()
    
    if request.user.is_staff:
        registros_pendentes = User.objects.filter(status='pendente').all()
        usuarios_aprovados = User.objects.filter(status='aprovado').all()
        context['registros'] = registros_pendentes
        context['usuarios_aprovados'] = usuarios_aprovados
    
    # KPI Data
    context['kpi_users_active'] = User.objects.filter(status='aprovado').count()
    context['kpi_pending_requests'] = User.objects.filter(status='pendente').count()
    context['kpi_traffic_alerts'] = SystemService.objects.filter(name='Congestionamento', status='offline').count()

    return render(request, 'portal_de_conexoes.html', context)

# Restrict to User ID 1
def is_super_admin_id(user):
    return user.is_superuser

@require_POST
@login_required
@user_passes_test(is_super_admin_id)
@csrf_exempt
def manage_service(request):
    try:
        data = json.loads(request.body)
        action = data.get('action')
        
        if action == 'add':
            service = SystemService.objects.create(name=data['name'], status=data['status'])
            return JsonResponse({'status': 'success', 'message': 'Serviço adicionado!', 'id': service.id})
            
        elif action == 'edit':
            service = get_object_or_404(SystemService, id=data['id'])
            service.name = data.get('name', service.name)
            service.status = data.get('status', service.status)
            service.save()
            return JsonResponse({'status': 'success', 'message': 'Serviço atualizado!'})
            
        elif action == 'delete':
            service = get_object_or_404(SystemService, id=data['id'])
            service.delete()
            return JsonResponse({'status': 'success', 'message': 'Serviço removido!'})
            
        return JsonResponse({'status': 'error', 'message': 'Ação inválida'}, status=400)
    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)}, status=500)

@require_POST
@login_required
@csrf_exempt
def manage_notification(request):
    try:
        data = json.loads(request.body)
        action = data.get('action', 'add')

        # Actions restricted to Super Admin
        if action in ['add', 'edit', 'delete']:
            if not is_super_admin_id(request.user):
                 return JsonResponse({'status': 'error', 'message': 'Permissão negada.'}, status=403)

            if action == 'add':
                GlobalNotification.objects.create(
                    title=data['title'],
                    message=data['message'],
                    created_by=request.user
                )
                return JsonResponse({'status': 'success', 'message': 'Notificação enviada!'})

            elif action == 'edit':
                notification = get_object_or_404(GlobalNotification, id=data['id'])
                notification.title = data.get('title', notification.title)
                notification.message = data.get('message', notification.message)
                notification.save()
                return JsonResponse({'status': 'success', 'message': 'Notificação atualizada!'})

            elif action == 'delete':
                notification = get_object_or_404(GlobalNotification, id=data['id'])
                notification.delete()
                return JsonResponse({'status': 'success', 'message': 'Notificação removida!'})

        # Actions available to all logged in users
        elif action == 'clear_all':
             notifications = GlobalNotification.objects.filter(is_active=True)
             for notif in notifications:
                 NotificationUserStatus.objects.update_or_create(
                     notification=notif,
                     user=request.user,
                     defaults={'is_cleared': True, 'is_read': True} 
                 )
             return JsonResponse({'status': 'success', 'message': 'Notificações limpas!'})

        elif action == 'mark_read':
             # Helper for potential future generic read marking
             notif_id = data.get('id')
             if notif_id:
                 notif = get_object_or_404(GlobalNotification, id=notif_id)
                 NotificationUserStatus.objects.update_or_create(
                     notification=notif,
                     user=request.user,
                     defaults={'is_read': True}
                 )
                 return JsonResponse({'status': 'success', 'message': 'Marcada como lida'})

        return JsonResponse({'status': 'error', 'message': 'Ação inválida'}, status=400)
    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)}, status=500)

def is_admin(user):
    return user.is_authenticated and user.is_superuser



@require_POST
@user_passes_test(is_admin)
def aprovar_registro(request, id):
    try:
        user = servicos.approve_user_service(id)
        messages.success(request, f'Usuário "{user.username}" aprovado com sucesso!')
    except ValueError:
        return HttpResponseBadRequest("Usuário Inválido")
    
    if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
        return JsonResponse({'status': 'success', 'message': f'Usuário "{user.username}" aprovado com sucesso!', 'username': user.username, 'id': user.id})

    return redirect(reverse('portal_de_conexoes'))

@require_POST
@user_passes_test(is_admin)
def rejeitar_registro(request, id):
    user = get_object_or_404(User.objects, id=id)
    username = user.username
    
    try:
        with transaction.atomic():
            # LogEntry, ActivityLog and Compromisso are automatically deleted via CASCADE
            user.delete()
        
        messages.warning(request, f'Registro de "{username}" rejeitado e removido.')
        
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
            return JsonResponse({'status': 'success', 'message': f'Registro de "{username}" rejeitado e removido.', 'action': 'deleted'})

    except Exception as e:
        messages.error(request, f'Erro ao remover "{username}": {e}')
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
            return JsonResponse({'status': 'error', 'message': f'Erro ao remover "{username}": {e}'}, status=500)
        
    return redirect(reverse('portal_de_conexoes'))

@require_POST
@user_passes_test(is_admin)
def remover_usuario(request, id):
    user = get_object_or_404(User.objects, id=id)
    username = user.username

    if user.is_staff or user.is_superuser:
        messages.error(request, 'Não é possível remover um administrador por este painel.')
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
            return JsonResponse({'status': 'error', 'message': 'Não é possível remover um administrador por este painel.'}, status=403)
        return redirect(reverse('portal_de_conexoes'))
    
    try:
        with transaction.atomic():
            # LogEntry, ActivityLog and Compromisso are automatically deleted via CASCADE
            user.delete()

        messages.error(request, f'Usuário "{username}" removido com sucesso.')
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
            return JsonResponse({'status': 'success', 'message': f'Usuário "{username}" removido com sucesso.', 'action': 'deleted'})
            
    except Exception as e:
        messages.error(request, f'Erro ao remover "{username}": {e}')
        if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
            return JsonResponse({'status': 'error', 'message': f'Erro ao remover "{username}": {e}'}, status=500)

    return redirect(reverse('portal_de_conexoes'))

@require_POST
@user_passes_test(is_admin)
def editar_usuario(request, id):
    user_to_edit = get_object_or_404(User.objects, id=id)
    
    if user_to_edit.is_superuser:
        messages.error(request, 'Não é possível editar os dados de um super-administrador.')
        return redirect(reverse('portal_de_conexoes'))

    new_username = request.POST.get('username')
    new_password = request.POST.get('password')

    if not new_username:
        messages.error(request, 'O nome de usuário não pode ser vazio.')
        return redirect(reverse('portal_de_conexoes'))

    if User.objects.filter(username=new_username).exclude(id=id).exists():
        messages.error(request, f'O nome de usuário "{new_username}" já está em uso.')
        return redirect(reverse('portal_de_conexoes'))

    user_to_edit.username = new_username
    if new_password:
        user_to_edit.set_password(new_password)
    
    user_to_edit.save()
    messages.success(request, f'Usuário "{new_username}" atualizado com sucesso!')
    
    if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
        return JsonResponse({'status': 'success', 'message': f'Usuário "{new_username}" atualizado com sucesso!', 'username': user_to_edit.username, 'id': user_to_edit.id})

    return redirect(reverse('portal_de_conexoes'))

@login_required
@require_GET
def get_user_profile(request):
    profile, created = UserProfile.objects.get_or_create(
        user=request.user, 
        defaults={'name': request.user.username, 'hire_date': date.today()}
    )
    
    image_url = None
    if profile.image_url:
        image_url = request.build_absolute_uri(profile.image_url.url)

    profile_data = {
        "name": profile.name or request.user.username,
        "role": "Administrador" if request.user.is_staff else "Usuário",
        "email": profile.email,
        "phone": profile.phone,
        "hireDate": profile.hire_date.strftime('%Y-%m-%d') if profile.hire_date else None,
        "bio": profile.bio,
        "imageUrl": image_url
    }
    return JsonResponse(profile_data)

@login_required
@csrf_exempt
def update_user_profile(request):
    if request.method != 'PUT':
        return HttpResponseBadRequest("Método inválido")

    profile, created = UserProfile.objects.get_or_create(user=request.user)
    
    profile.name = request.POST.get('name', profile.name)
    profile.email = request.POST.get('email', profile.email)
    profile.phone = request.POST.get('phone', profile.phone)
    profile.bio = request.POST.get('bio', profile.bio)
    
    if 'profileImage' in request.FILES:
        profile.image_url = request.FILES['profileImage']

    try:
        profile.save()
        
        image_url_resp = None
        if profile.image_url:
            image_url_resp = request.build_absolute_uri(profile.image_url.url)
            
        return JsonResponse({
            "message": "Perfil atualizado com sucesso!", 
            "name": profile.name,
            "role": "Administrador" if request.user.is_staff else "Usuário", 
            "imageUrl": image_url_resp
        }, status=200)
    except Exception as e:
        return JsonResponse({"message": "Ocorreu um erro ao salvar as alterações."}, status=500)



@login_required
@csrf_exempt
def compromisso_list_create_view(request):
    if request.method == 'GET':
        compromissos = Compromisso.objects.filter(user_id=request.user.id).order_by('date', 'start_time')
        data = [c.to_dict() for c in compromissos]
        return JsonResponse({"compromissos": data})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            print(f"DEBUG: Saving Event Data: {data}")
            
            end_time = None
            if data.get('end_time'):
                end_time = datetime.strptime(data['end_time'], '%H:%M').time()

            new_compromisso = Compromisso.objects.create(
                title=data['title'],
                date=datetime.strptime(data['date'], '%Y-%m-%d').date(),
                start_time=datetime.strptime(data['start_time'], '%H:%M').time(),
                end_time=end_time,
                description=data.get('description'),
                category=data.get('category'),
                status=data.get('status', 'Agendado'),
                user=request.user
            )
            return JsonResponse(new_compromisso.to_dict(), status=201)
        except (KeyError, ValueError) as e:
            print(f"DEBUG: Value Error saving Event: {e}")
            return JsonResponse({"message": f"Dados inválidos: {str(e)}"}, status=400)
        except Exception as e:
            print(f"DEBUG: Internal Error saving Event: {e}")
            return JsonResponse({"message": f"Erro interno: {str(e)}"}, status=500)
    
    return HttpResponseBadRequest("Método não suportado")

@login_required
@csrf_exempt
def compromisso_detail_view(request, compromisso_id):
    compromisso = get_object_or_404(Compromisso.objects, id=compromisso_id)
    if compromisso.user_id != request.user.id:
        return HttpResponseForbidden()

    if request.method == 'GET':
        return JsonResponse(compromisso.to_dict())

    if request.method == 'PUT':
        try:
            data = json.loads(request.body)
            compromisso.title = data.get('title', compromisso.title)
            compromisso.date = datetime.strptime(data['date'], '%Y-%m-%d').date() if data.get('date') else compromisso.date
            compromisso.start_time = datetime.strptime(data['start_time'], '%H:%M').time() if data.get('start_time') else compromisso.start_time
            compromisso.end_time = datetime.strptime(data['end_time'], '%H:%M').time() if data.get('end_time') else None
            compromisso.description = data.get('description', compromisso.description)
            compromisso.category = data.get('category', compromisso.category)
            compromisso.status = data.get('status', compromisso.status)
            compromisso.save()
            return JsonResponse(compromisso.to_dict())
        except (ValueError):
            return JsonResponse({"message": "Formato de data ou hora inválido."}, status=400)
        except Exception as e:
            return JsonResponse({"message": f"Erro interno: {str(e)}"}, status=500)

    if request.method == 'DELETE':
        try:
            compromisso.delete()
            return JsonResponse({"message": "Compromisso excluído com sucesso."}, status=200)
        except Exception as e:
            return JsonResponse({"message": f"Erro interno: {str(e)}"}, status=500)

    return HttpResponseBadRequest("Método não suportado")

@login_required
@require_GET
def get_users_list(request):
    try:
        users = User.objects.filter(status='aprovado').exclude(id=request.user.id)
        user_list = [{"id": user.id, "name": user.username} for user in users]
        return JsonResponse({"users": user_list})
    except Exception as e:
        return JsonResponse({"message": f"Erro interno: {str(e)}"}, status=500)

@login_required
@require_POST
@csrf_exempt
def share_compromisso_view(request, compromisso_id):
    compromisso_original = get_object_or_404(Compromisso.objects, id=compromisso_id)
    if compromisso_original.user_id != request.user.id:
        return HttpResponseForbidden("Você não tem permissão para compartilhar este compromisso.")
    
    try:
        data = json.loads(request.body)
        user_ids_to_share = data.get('user_ids')
    except json.JSONDecodeError:
        return JsonResponse({"message": "Requisição inválida."}, status=400)

    if not user_ids_to_share or not isinstance(user_ids_to_share, list):
        return JsonResponse({"message": "A lista de IDs de usuários é inválida."}, status=400)

    try:
        with transaction.atomic():
            for user_id in user_ids_to_share:
                Compromisso.objects.create(
                    title=f"[Compartilhado de {request.user.username}] {compromisso_original.title}",
                    date=compromisso_original.date,
                    start_time=compromisso_original.start_time,
                    end_time=compromisso_original.end_time,
                    description=compromisso_original.description,
                    category=compromisso_original.category,
                    status=compromisso_original.status,
                    user_id=int(user_id) # Keeping as ID since we pass ID, but Django handles it fine for FK too if named user_id
                )
        return JsonResponse({"message": "Compromisso compartilhado com sucesso!"}, status=200)
    except Exception as e:
        return JsonResponse({"message": f"Erro interno: {str(e)}"}, status=500)

@csrf_exempt
@require_POST
def hannah_chat_view(request):
    try:
        data = json.loads(request.body)
        user_message = data.get('message')
    except json.JSONDecodeError:
        return JsonResponse({'reply': 'Ocorreu um erro. Nenhuma mensagem recebida.'}, status=400)

    if not user_message:
        return JsonResponse({'reply': 'Ocorreu um erro. Nenhuma mensagem recebida.'}, status=400)

    username = request.user.username if request.user.is_authenticated else "Visitante"
    hannah_response = get_hannah_response(user_message, username)
    
    # Check if the response is a dictionary (new format) or string (old format fallback)
    if isinstance(hannah_response, dict):
        return JsonResponse({
            'reply': hannah_response['text'],
            'action': hannah_response.get('action')
        })
    else:
        # Fallback handling
        return JsonResponse({'reply': hannah_response})

@login_required
@require_GET
def get_latest_notification(request):
    try:
        latest_notification = GlobalNotification.objects.filter(is_active=True).order_by('-created_at').first()
        latest_id = latest_notification.id if latest_notification else 0
        return JsonResponse({"latest_id": latest_id})
    except Exception as e:
        return JsonResponse({"message": f"Erro interno: {str(e)}"}, status=500)