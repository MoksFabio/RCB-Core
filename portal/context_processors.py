from .models import GlobalNotification, NotificationUserStatus

def is_super_admin(request):
    context = {}
    if request.user.is_authenticated:
        context['is_super_admin'] = request.user.is_superuser
        
        # Add Global Notifications
        # Filter out notifications cleared by this user
        cleared_ids = NotificationUserStatus.objects.filter(
            user=request.user, 
            is_cleared=True
        ).values_list('notification_id', flat=True)
        
        notifications = GlobalNotification.objects.filter(
            is_active=True
        ).exclude(
            id__in=cleared_ids
        ).order_by('-created_at')
        
        
        # Calculate unread count and get read IDs
        read_ids = list(NotificationUserStatus.objects.filter(
            user=request.user,
            is_read=True
        ).values_list('notification_id', flat=True))
        
        unread_count = notifications.exclude(id__in=read_ids).count()
        

        
        context['notifications'] = notifications
        context['unread_notification_count'] = unread_count
        context['read_notification_ids'] = read_ids
        
        if notifications.exists():
            context['latest_notification_id'] = notifications.first().id
        else:
            context['latest_notification_id'] = 0
    else:
        context['is_super_admin'] = False
        context['latest_notification_id'] = 0
        
    return context
