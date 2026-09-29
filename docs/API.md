# API

Базовый путь — `/api`. Интерактивная документация: `/api/docs`, схема OpenAPI: `/api/openapi.json` (копия лежит в `docs/openapi.json`).

Авторизация через cookie `lukas_session`: её ставят регистрация и вход. Все методы, кроме регистрации, входа, `/api/system` и `/api/health`, требуют входа.

## Обычный сценарий

```bash
API=http://localhost:8001/api

# 1. Регистрация (cookie сохранится в jar.txt)
curl -c jar.txt -H 'Content-Type: application/json' \
  -d '{"email":"me@example.com","password":"password123"}' $API/auth/register

# 2. Рабочее пространство пользователя
PROJECT=$(curl -s -b jar.txt $API/workspace | python -c "import sys,json;print(json.load(sys.stdin)['id'])")

# 3. Загрузка шаблона — сразу запускается его разбор
TEMPLATE=$(curl -s -b jar.txt -F file=@template.pptx $API/templates | python -c "import sys,json;print(json.load(sys.stdin)['id'])")
curl -b jar.txt $API/templates/$TEMPLATE      # ждём preparation.status == "ready"

# 4. Запуск генерации
curl -b jar.txt -H 'Content-Type: application/json' \
  -d "{\"template_id\":\"$TEMPLATE\",\"brief\":\"О чём презентация\",\"slide_count\":8}" \
  $API/projects/$PROJECT/runs

# 5. Статус запуска: status, stage, progress; по готовности versions — три варианта
curl -b jar.txt $API/runs/<run_id>

# 6. Вариант и его файлы (pptx, pdf, html, превью, отчёт аудита)
curl -b jar.txt $API/versions/<version_id>
curl -b jar.txt -OJ $API/artifacts/<artifact_id>
```

## Методы

### Аккаунт

| Метод | Путь | Что делает |
| --- | --- | --- |
| POST | `/auth/register` | регистрация: `email`, `password` (от 8 символов) |
| POST | `/auth/login` | вход |
| GET | `/auth/me` | текущий пользователь |
| POST | `/auth/logout` | выход |
| GET | `/workspace` | проект пользователя (создаётся при регистрации) |

### Шаблоны (общая библиотека)

| Метод | Путь | Что делает |
| --- | --- | --- |
| GET | `/templates` | список шаблонов |
| POST | `/templates` | загрузить PPTX (`file`), запускает разбор |
| GET | `/templates/{id}` | шаблон и статус разбора |
| POST | `/templates/{id}/prepare` | повторить разбор после ошибки |
| GET | `/templates/{id}/previews/{n}` | картинка n-го слайда шаблона |
| GET | `/templates/{id}/export` | скачать шаблон с готовым разбором (.zip) |
| POST | `/templates/import` | загрузить такой .zip без платного разбора |

### Проекты и материалы

| Метод | Путь | Что делает |
| --- | --- | --- |
| GET / POST | `/projects` | список / создание проекта |
| GET | `/projects/{id}` | проект |
| POST | `/projects/{id}/sources` | загрузить файл: `file` и `kind` = `content` (TXT, MD, CSV, JSON, PDF, DOCX) или `template` (PPTX) |
| GET | `/projects/{id}/sources` | файлы проекта |
| DELETE | `/projects/{id}/sources/{source_id}` | убрать шаблон из проекта |
| GET | `/sources/{id}` | файл и статус разбора |

### Генерация

| Метод | Путь | Что делает |
| --- | --- | --- |
| POST | `/projects/{id}/runs` | запуск: `brief`, `template_id`, `slide_count` (1–60, по умолчанию 10), `content_source_ids` |
| GET | `/projects/{id}/runs` | история запусков |
| GET | `/runs/{id}` | статус и результат |
| POST | `/runs/{id}/cancel` | остановить |
| POST | `/runs/{id}/retry` | повторить упавший запуск |

Заголовок `Idempotency-Key` защищает от двойного запуска при повторной отправке.

### Варианты, аудит и правки

| Метод | Путь | Что делает |
| --- | --- | --- |
| GET | `/projects/{id}/versions` | все версии проекта |
| GET | `/versions/{id}` | версия: `variant_id` (`a` — базовая, `b` — больше текста, `c` — больше визуала), файлы, превью |
| GET | `/versions/{id}/issues` | замечания аудита |
| POST | `/versions/{id}/repairs` | исправить выбранные замечания: `issue_ids` → новая версия |
| POST | `/versions/{id}/edits` | переписать слайд по просьбе: `slide_index`, `prompt` → новая версия |
| GET | `/artifacts/{id}` | скачать файл |

### Служебные

| Метод | Путь | Что делает |
| --- | --- | --- |
| GET | `/health` | проверка, что API и база живы |
| GET | `/system` | какие модели настроены и есть ли ключ (сам ключ не отдаётся) |
