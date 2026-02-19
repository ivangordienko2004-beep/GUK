from django.conf import settings
from django.shortcuts import redirect
from django.urls import reverse

from .auth_utils import DEFAULT_ADMIN_PASSWORD, DEFAULT_ADMIN_USERNAME, ensure_default_users


class ForceAdminPasswordChangeMiddleware:
    """Force default admin account to change password after login."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        ensure_default_users()

        if request.user.is_authenticated and request.user.username == DEFAULT_ADMIN_USERNAME:
            using_default_password = request.user.check_password(DEFAULT_ADMIN_PASSWORD)
            if using_default_password:
                change_url = reverse('password_change')
                allowed_paths = {
                    change_url,
                    reverse('password_change_done'),
                    reverse('logout'),
                }
                if request.path not in allowed_paths and not request.path.startswith(settings.STATIC_URL):
                    return redirect('password_change')

        return self.get_response(request)
