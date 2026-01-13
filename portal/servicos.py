from django.db import transaction
from .models import User, UserProfile
from datetime import date
from django.core.exceptions import ObjectDoesNotExist

def create_user_service(username, password):
    """
    Creates a new user and their profile within a transaction.
    Raises ValueError if username already exists.
    """
    if User.objects.filter(username=username).exists():
        raise ValueError("Este nome de usuário já está em uso.")

    with transaction.atomic():
        user = User(username=username)
        user.set_password(password)
        user.save()
        
        UserProfile.objects.create(
            user=user, 
            name=username, 
            hire_date=date.today()
        )
    return user

def approve_user_service(user_id):
    """
    Approves a pending user.
    """
    try:
        user = User.objects.get(id=user_id)
        user.status = 'aprovado'
        user.save()
        return user
    except ObjectDoesNotExist:
        raise ValueError("Usuário não encontrado.")
