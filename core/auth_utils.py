from django.contrib.auth import get_user_model
from django.db.utils import OperationalError, ProgrammingError


DEFAULT_ADMIN_USERNAME = 'admin'
DEFAULT_ADMIN_PASSWORD = 'admin'
DEFAULT_GUEST_USERNAME = 'guest'
DEFAULT_GUEST_PASSWORD = 'guest'


def ensure_default_users() -> None:
    """Create default admin/guest users on first run if they do not exist yet."""
    user_model = get_user_model()

    try:
        admin, admin_created = user_model.objects.get_or_create(
            username=DEFAULT_ADMIN_USERNAME,
            defaults={
                'is_staff': True,
                'is_superuser': True,
                'is_active': True,
            },
        )
        if admin_created:
            admin.set_password(DEFAULT_ADMIN_PASSWORD)
            admin.save(update_fields=['password'])

        guest, guest_created = user_model.objects.get_or_create(
            username=DEFAULT_GUEST_USERNAME,
            defaults={
                'is_staff': False,
                'is_superuser': False,
                'is_active': True,
            },
        )
        if guest_created:
            guest.set_password(DEFAULT_GUEST_PASSWORD)
            guest.save(update_fields=['password'])
    except (OperationalError, ProgrammingError):
        # DB might not be migrated yet.
        return
