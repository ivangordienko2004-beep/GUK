# GUK — Django приложение для ВУЦ

Веб-приложение с двумя уровнями доступа:
- **пользователь**: загрузка Excel, объединение, скачивание, создание отчёта;
- **администратор** (`is_staff=True`): всё выше + кнопка **«Расшифровка»**.

## Возможности

1. Авторизация (логин/пароль).
2. Загрузка нескольких Excel-файлов (`.xlsx/.xls`).
3. Объединение в одну таблицу и скачивание.
4. Расшифровка по номеру ВУС (только админ):
   - для офицеров поля должности очищаются;
   - для сержантов по номеру должности подставляется наименование.
5. Формирование итогового отчёта в отдельный Excel-файл (другой формат).
6. Интерфейс в military-style цветах.

## Локальный запуск без контейнера

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

После входа обычный пользователь также может работать через `/login/`.

## Запуск в Docker-контейнере на базе Astra Linux

Приложение упаковано в Docker-образ на базе официального образа Astra Linux Special Edition 1.8 с Python 3.11:

```text
registry.astralinux.ru/library/astra/ubi18-python311:latest
```

### Быстрый старт через Docker Compose v2

1. Подготовьте переменные окружения:

   ```bash
   cp .env.example .env
   # обязательно замените DJANGO_SECRET_KEY в .env на длинное случайное значение
   ```

2. Соберите и запустите контейнер:

   ```bash
   docker compose up --build -d
   ```

3. Создайте администратора:

   ```bash
   docker compose exec guk python manage.py createsuperuser
   ```

4. Откройте приложение:

   ```text
   http://localhost:8000/
   ```

При старте контейнер автоматически выполняет миграции и собирает статику.
Данные SQLite, загруженные файлы и собранная статика сохраняются в именованных томах `guk_data`, `guk_media` и `guk_static`.

### Ручная сборка и запуск Docker

```bash
docker build -t guk:astra .
docker run --rm \
  --name guk-astra \
  -p 8000:8000 \
  -e DJANGO_SECRET_KEY='replace-with-a-long-random-secret' \
  -e DJANGO_ALLOWED_HOSTS='localhost,127.0.0.1' \
  -v guk_data:/app/data \
  -v guk_media:/app/media \
  -v guk_static:/app/staticfiles \
  guk:astra
```

## Конфигурация контейнера

| Переменная | Значение по умолчанию в контейнере | Назначение |
| --- | --- | --- |
| `DJANGO_SECRET_KEY` | `change-me-in-production` в Compose | Секретный ключ Django. Для боевого запуска обязательно заменить. |
| `DJANGO_DEBUG` | `false` | Включение/выключение debug-режима. |
| `DJANGO_ALLOWED_HOSTS` | `localhost,127.0.0.1` | Список разрешенных хостов через запятую. |
| `DJANGO_CSRF_TRUSTED_ORIGINS` | пусто | Доверенные origin для CSRF, например `https://guk.example.ru`. |
| `DJANGO_STATIC_ROOT` | `/app/staticfiles` | Каталог собранной статики внутри контейнера. |
| `DJANGO_MEDIA_ROOT` | `/app/media` | Каталог пользовательских файлов внутри контейнера. |
| `GUK_DATABASE_NAME` | `/app/data/db.sqlite3` | Путь к SQLite-базе внутри контейнера. |

## Проверка пакета

Минимальные проверки перед публикацией образа:

```bash
python manage.py check
python -m compileall -q .
docker compose build
docker compose up -d
docker compose exec guk python manage.py check
docker compose down
```
