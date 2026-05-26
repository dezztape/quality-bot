# 🚀 Deployment Guide - Quality Bot на Railway

## Подготовка

### 1. Подготовить репозиторий

```bash
# .gitignore должен содержать:
echo ".env" >> .gitignore
echo "__pycache__/" >> .gitignore
echo "*.pyc" >> .gitignore
echo "storage/" >> .gitignore
echo "*.xlsx" >> .gitignore
```

### 2. Убедиться что requirements.txt актуален

```bash
pip freeze > requirements.txt
```

### 3. Создать Procfile для Railway

```bash
# Создать файл Procfile в корне проекта:
echo "worker: python main.py" > Procfile
```

## Развертывание на Railway

### Шаг 1: Создать проект на Railway

1. Перейти на https://railway.app
2. Нажать "New Project"
3. Выбрать "Deploy from GitHub"
4. Подключить репозиторий

### Шаг 2: Добавить PostgreSQL

1. В Railway: "Add Service" → PostgreSQL
2. Railway автоматически создаст `RAILWAY_DATABASE_URL`

### Шаг 3: Конфигурировать переменные

В Railway > Variables добавить:

```
BOT_TOKEN = <ваш токен от @BotFather>
DATABASE_URL = $DATABASE_URL  (автоматически от PostgreSQL)
ADMIN_IDS = 123456789,987654321
REDIS_URL = <оставить пусто, может быть добавлено позже>
```

### Шаг 4: Запустить

1. Railway автоматически обнаружит Procfile
2. Бот начнет работать с polling
3. Проверить логи в Railway > Logs

## 🔐 Безопасность на Railway

### ✅ Уже защищено:

- DATABASE_URL никогда не виден в коде
- PostgreSQL использует SSL-соединение
- BOT_TOKEN хранится в переменных окружения

### ⚠️ Дополнительные меры (опционально):

#### Добавить Redis для лучшей производительности:

```
1. Railway: "Add Service" → Redis
2. Railway создаст RAILWAY_REDIS_URL
3. Скопировать в REDIS_URL переменную
```

#### Включить Webhooks (вместо polling):

- Дешевле чем polling
- Более безопасно
- Требует фиксированного адреса

## 📊 Мониторинг

### Логирование

Все важные события логируются:

- Старт пользователя
- Начало/завершение тестов
- Ошибки базы данных
- Подключение Redis (если используется)

### Просмотр логов

```bash
railway logs -f  # в терминале (если установлен Railway CLI)
# или через Railway > Logs в веб-интерфейсе
```

### Типичные логи:

```
2026-05-26 14:23:45 - __main__ - INFO - 🤖 Quality Bot starting...
2026-05-26 14:23:45 - __main__ - INFO - ✅ Loaded 2 admin(s)
2026-05-26 14:23:47 - __main__ - INFO - 🚀 Bot started, polling messages...
2026-05-26 14:25:10 - __main__ - INFO - 👤 User 123456789 started bot
2026-05-26 14:25:20 - __main__ - INFO - 📝 User 123456789 starting control test
2026-05-26 14:30:45 - __main__ - INFO - ✅ User 123456789 completed test 1 (control) in 2 attempt(s)
```

## 🔄 Обновление кода

```bash
# 1. Обновить код локально
# 2. Коммитить в GitHub
# 3. Railway автоматически переразвернет бот

git add .
git commit -m "Update: add new features"
git push origin main
# Railway автоматически увидит обновление и переразвернет
```

## 💾 Резервная копия базы данных

Railway автоматически хранит резервные копии PostgreSQL.
Для ручного экспорта:

```bash
# Локально установить pgAdmin или psql
psql $RAILWAY_DATABASE_URL -c "\dt"  # список таблиц
```

## 🆘 Troubleshooting

### Бот не отвечает

- Проверить логи: `railway logs -f`
- Проверить DATABASE_URL установлен правильно
- Убедиться что BOT_TOKEN правильный

### Ошибка подключения к БД

```
sqlalchemy.exc.OperationalError: (asyncpg.exceptions.CannotConnectNowError)
```

- Может быть Railway перезагружается
- Проверить статус сервиса в Railway dashboard
- Переразвернуть: Railway > Redeploy

### Высокое использование памяти

- MemoryStorage накапливает данные в памяти
- Решение: добавить Redis (используется автоматически если REDIS_URL установлен)

## 📝 .gitignore для Railway

```
# Окружение
.env
.env.local
*.pem

# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
venv/
env/

# IDE
.vscode/
.idea/
*.swp

# Проект
storage/
*.xlsx
*.csv
results_*
```

## ✨ Production Checklist

- [ ] BOT_TOKEN установлен в Railway
- [ ] DATABASE_URL установлен в Railway
- [ ] ADMIN_IDS설정
- [ ] Логи проверены на ошибки
- [ ] Тестовый юзер успешно зарегистрирован
- [ ] Админ может создавать и утверждать тесты
- [ ] Обычный юзер может проходить утвержденные тесты
- [ ] Экспорт CSV/Excel работает
- [ ] Статистика отображается корректно
- [ ] (Опционально) Redis подключен для масштабирования

## 🚨 Emergency

Если нужно временно отключить бот:

- Railway > Undeploy
- Бот прекратит работу (~30 сек)
- Railway > Redeploy для включения
