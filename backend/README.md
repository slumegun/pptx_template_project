# Бэкенд

FastAPI-сервер (`app/`) и движок генерации (`engine/`). Как запустить весь проект — в корневом [README](../README.md).

## Запуск

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -e ".[test]"
uvicorn app.main:app --reload --port 8001
```

Настройки берутся из `backend/.env`. По умолчанию там SQLite и очередь внутри процесса (`AYA_LOCAL_JOBS=true`), поэтому Redis не нужен. Для генерации впишите `OPENROUTER_API_KEY`.

В Docker тот же код работает как API (`uvicorn`), воркер (`python -m app.jobs`) и отдельный сервис миграций (`alembic upgrade head`).

## Где что

- `app/main.py` — все маршруты API, `app/auth.py` — аккаунты, `app/library.py` — библиотека шаблонов.
- `app/jobs.py` — выполнение задач: разбор шаблона, генерация, ремонт, правка слайда.
- `engine/pipeline.py` — порядок шагов генерации.
- `engine/ingest.py` — разбор шаблона, `engine/renderer.py` — выбор макетов и сборка PPTX.
- `engine/visuals.py` — поиск чисел для графиков и плиток.
- `engine/provider.py` — запросы к OpenRouter, `engine/prompts/` — промпты агентов.
- `engine/audit.py`, `design_audit.py`, `rendered_audit.py` — проверки готовых файлов.

## Тесты и скрипты

```bash
pytest -q                                # без сети и без трат
python scripts/check_models.py           # проверить настройки моделей
python scripts/check_models.py --live    # платно: по запросу в каждую модель
python scripts/smoke_api.py --live       # платно: полный прогон через API на примере из examples/
```

## Миграции

```bash
alembic upgrade head
alembic revision --autogenerate -m "что изменилось"
```
