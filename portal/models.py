from django.db import models
from django.contrib.auth.models import AbstractUser
from django.conf import settings
from django.utils import timezone
from datetime import date
import uuid

def generate_manifestacao_id():
    return "PROC-" + str(uuid.uuid4())[:8].upper()

class User(AbstractUser):
    STATUS_CHOICES = [
        ('pendente', 'Pendente'),
        ('aprovado', 'Aprovado'),
    ]
    
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pendente')

class UserProfile(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='profile')
    name = models.CharField(max_length=150, blank=True, null=True)
    email = models.EmailField(max_length=150, blank=True, null=True)
    phone = models.CharField(max_length=20, blank=True, null=True)
    hire_date = models.DateField(blank=True, null=True)
    bio = models.TextField(blank=True, null=True)
    image_url = models.ImageField(upload_to='profile_images/', max_length=256, blank=True, null=True)

    def __str__(self):
        return self.user.username

class Compromisso(models.Model):
    id = models.AutoField(primary_key=True)
    title = models.CharField(max_length=200)
    date = models.DateField()
    start_time = models.TimeField()
    end_time = models.TimeField(blank=True, null=True)
    description = models.TextField(blank=True, null=True)
    category = models.CharField(max_length=100, blank=True, null=True)
    status = models.CharField(max_length=50, default='Agendado')
    
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='compromissos', null=True)

    def to_dict(self):
        return {
            "id": self.id,
            "title": self.title,
            "date": self.date.strftime('%Y-%m-%d') if self.date else None,
            "start_time": self.start_time.strftime('%H:%M') if self.start_time else None,
            "end_time": self.end_time.strftime('%H:%M') if self.end_time else None,
            "description": self.description,
            "category": self.category,
            "status": self.status,
            "user_id": self.user_id
        }
    
    class Meta:
        db_table = 'compromisso'

class Manifestacao(models.Model):
    id = models.CharField(max_length=40, primary_key=True, default=generate_manifestacao_id)
    numero_processo = models.CharField(max_length=100, unique=True)
    data_entrada = models.DateField()
    prazo_final = models.DateField()
    descricao = models.TextField()
    situacao = models.CharField(max_length=50, default="PENDENTE")
    atribuicao = models.CharField(max_length=150, blank=True, null=True)
    envio = models.CharField(max_length=200, blank=True, null=True)
    check_visual = models.CharField(max_length=20, blank=True, null=True)

    @property
    def dias_restantes(self):
        if not self.prazo_final:
            return None
        
        today_date_obj = timezone.now().date()
        remaining_days = (self.prazo_final - today_date_obj).days
        return remaining_days

    def to_dict(self):
        return {
            "id": self.id,
            "Nº do Processo": self.numero_processo,
            "Data de Entrada": self.data_entrada.strftime('%Y-%m-%d') if self.data_entrada else None,
            "Prazo Final": self.prazo_final.strftime('%Y-%m-%d') if self.prazo_final else None,
            "Descrição": self.descricao,
            "Situação": self.situacao,
            "Atribuição": self.atribuicao,
            "Envio": self.envio,
            "Check": self.check_visual,
            "Dias Restantes": self.dias_restantes
        }

    class Meta:
        db_table = 'manifestacao'


class ActivityLog(models.Model):
    id = models.AutoField(primary_key=True)
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='activity_logs', null=True)
    user_name = models.CharField(max_length=150)
    message = models.CharField(max_length=500)
    status = models.CharField(max_length=50, default='info')
    timestamp = models.DateTimeField(default=timezone.now)

    def to_dict(self):
        return {
            'message': self.message,
            'status': self.status,
            'timestamp': self.timestamp.isoformat() + 'Z'
        }

    class Meta:
        db_table = 'activity_log'

class SystemService(models.Model):
    STATUS_CHOICES = [
        ('operando', 'Operando'),
        ('instavel', 'Instável'),
        ('offline', 'Offline'),
    ]
    
    name = models.CharField(max_length=100)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='operando')
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.name} - {self.status}"

    class Meta:
        db_table = 'system_service'
        ordering = ['name']

class GlobalNotification(models.Model):
    title = models.CharField(max_length=200)
    message = models.TextField()
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.title

    class Meta:
        db_table = 'global_notification'
        ordering = ['-created_at']

class NotificationUserStatus(models.Model):
    notification = models.ForeignKey(GlobalNotification, on_delete=models.CASCADE, related_name='user_statuses')
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='notification_statuses')
    is_read = models.BooleanField(default=False)
    is_cleared = models.BooleanField(default=False)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'notification_user_status'
        unique_together = ['notification', 'user']