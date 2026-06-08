#!/usr/bin/env sh
set -eu

mkdir -p "${DJANGO_STATIC_ROOT:-/app/staticfiles}" "${DJANGO_MEDIA_ROOT:-/app/media}" "$(dirname "${GUK_DATABASE_NAME:-/app/data/db.sqlite3}")"

python manage.py migrate --noinput
python manage.py collectstatic --noinput

exec "$@"
