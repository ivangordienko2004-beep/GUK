from django.conf import settings
from django.shortcuts import redirect
from django.urls import NoReverseMatch, reverse

from .auth_utils import DEFAULT_ADMIN_PASSWORD, DEFAULT_ADMIN_USERNAME, ensure_default_users


def _safe_reverse(name: str) -> str | None:
    try:
        return reverse(name)
    except NoReverseMatch:
        return None


class ForceAdminPasswordChangeMiddleware:
    """Force default admin account to change password after login."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        ensure_default_users()

        if request.user.is_authenticated and request.user.username == DEFAULT_ADMIN_USERNAME:
            using_default_password = request.user.check_password(DEFAULT_ADMIN_PASSWORD)
            if using_default_password:
                change_url = _safe_reverse('password_change')
                done_url = _safe_reverse('password_change_done')
                logout_url = _safe_reverse('logout')

                if not change_url:
                    return self.get_response(request)

                allowed_paths = {change_url}
                if done_url:
                    allowed_paths.add(done_url)
                if logout_url:
                    allowed_paths.add(logout_url)

                if request.path not in allowed_paths and not request.path.startswith(settings.STATIC_URL):
                    return redirect(change_url)

        return self.get_response(request)
