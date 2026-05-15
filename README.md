# 52-я стрелковая дивизия — архив и боевой путь

Исторический сайт с поиском по людям, событиям и боевому пути дивизии.

## Стек

| Компонент   | Технология              |
|-------------|-------------------------|
| База данных | PostgreSQL 16           |
| API         | PostgREST v12           |
| Веб-сервер  | Nginx                   |
| Фронтенд    | Vanilla JS + Leaflet.js |
| Деплой      | Docker Compose          |

## Структура репозитория

```
division52/
├── docker-compose.yml        — все сервисы
├── .env.example              — шаблон переменных окружения
├── nginx/
│   └── nginx.conf            — конфиг Nginx
├── sql/
│   └── migrations/
│       └── 001_initial.sql   — схема БД
├── import/
│   └── import_word.py        — скрипт парсинга Word-файлов
└── frontend/
    ├── index.html
    ├── css/
    └── js/
```

## Быстрый старт на сервере (reg.ru)

### 1. Установка Docker

```bash
curl -fsSL https://get.docker.com | sh
apt install docker-compose-plugin -y
```

### 2. Клонирование репозитория

```bash
git clone https://github.com/ТВОЙаккаунт/division52.git
cd division52
```

### 3. Настройка окружения

```bash
cp .env.example .env
nano .env   # заполни пароли и домен
```

### 4. SSL сертификат (Let's Encrypt)

```bash
apt install certbot -y
certbot certonly --standalone -d твой-домен.ru
```

### 5. Папка для фотографий

```bash
mkdir -p /var/division52/photos/officers
```

### 6. Запуск

```bash
docker compose up -d
```

Проверка:
```bash
docker compose ps          # все сервисы должны быть healthy
curl localhost:3000/events  # PostgREST отвечает
```

---

## Импорт данных из Word

### Установка зависимостей

```bash
cd import
pip install python-docx anthropic psycopg2-binary python-dotenv
```

### Добавь ANTHROPIC_API_KEY в .env

```
ANTHROPIC_API_KEY=sk-ant-...
```

### Запуск импорта

```bash
# Сначала проверь без записи в БД
python import_word.py --file ../data/commanders.docx --type persons --dry-run

# Затем реальный импорт
python import_word.py --file ../data/commanders.docx --type persons
python import_word.py --file ../data/combat_path.docx --type events
```

---

## API — примеры запросов

PostgREST автоматически создаёт REST API из таблиц.

```bash
# Все события
GET /api/events

# Поиск по тексту
GET /api/events?search_vector=fts(russian).Курская+дуга

# События за 1943 год
GET /api/events?date_from=gte.1943-01-01&date_from=lte.1943-12-31

# Ключевые сражения
GET /api/events?is_key=eq.true&order=date_from

# Командиры конкретного полка
GET /api/service_records?unit_id=eq.3&select=*,persons(*)

# Человек со всей историей службы
GET /api/persons?id=eq.42&select=*,service_records(*,units(*))
```

---

## Загрузка фотографий

```bash
# Копируй фото на сервер
scp officers/*.jpg root@твой-домен.ru:/var/division52/photos/officers/

# Конвертация в webp для экономии трафика (опционально)
apt install webp -y
for f in /var/division52/photos/officers/*.jpg; do
  cwebp -q 85 "$f" -o "${f%.jpg}.webp"
done
```

Затем в БД:
```sql
UPDATE persons SET photo_url = '/photos/officers/ivanov-np.webp'
WHERE last_name = 'Иванов' AND first_name = 'Николай';
```

---

## Обновление кода

```bash
git pull
docker compose restart nginx  # для изменений фронтенда
docker compose restart postgrest  # если менялась схема БД
```

Для изменений схемы БД:
```bash
docker compose exec db psql -U $POSTGRES_USER -d $POSTGRES_DB \
  -f /docker-entrypoint-initdb.d/002_новая_миграция.sql
```
