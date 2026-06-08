# syntax=docker/dockerfile:1

FROM registry.astralinux.ru/library/astra/ubi18-python311:latest

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1 \
    DJANGO_DEBUG=false \
    DJANGO_ALLOWED_HOSTS=localhost,127.0.0.1 \
    DJANGO_STATIC_ROOT=/app/staticfiles \
    DJANGO_MEDIA_ROOT=/app/media \
    GUK_DATABASE_NAME=/app/data/db.sqlite3

WORKDIR /app

RUN groupadd --system guk \
    && useradd --system --gid guk --home-dir /app --shell /usr/sbin/nologin guk

COPY requirements.txt ./
RUN python -m pip install --upgrade pip \
    && python -m pip install -r requirements.txt

COPY --chown=guk:guk . .
RUN mkdir -p /app/staticfiles /app/media /app/data \
    && chown -R guk:guk /app

USER guk

EXPOSE 8000

ENTRYPOINT ["/app/docker/entrypoint.sh"]
CMD ["gunicorn", "guk_project.wsgi:application", "--bind", "0.0.0.0:8000", "--workers", "3"]
