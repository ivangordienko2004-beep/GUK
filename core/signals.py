from django.db.models.signals import post_migrate
from django.dispatch import receiver

from .auth_utils import ensure_default_users


@receiver(post_migrate)
def create_default_users_after_migrate(sender, **kwargs):
    """Ensure default users exist right after migrations complete."""
    ensure_default_users()
