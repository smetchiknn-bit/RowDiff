# Excel в Google Таблицы

Приложение для конвертации Excel файлов (.xlsx) в Google Таблицы с сохранением в выбранную папку на Google Диске.

## 🚀 Быстрый старт

### 1. Установка зависимостей
```bash
npm install
```

### 2. Настройка Google API (обязательно!)

#### Шаг 1: Создание проекта в Google Cloud Console
1. Перейдите в [Google Cloud Console](https://console.cloud.google.com/)
2. Нажмите **"Create Project"** (Создать проект)
3. Введите имя проекта (например, "Excel to Sheets")
4. Нажмите **"Create"**

#### Шаг 2: Включение необходимых API
1. В созданном проекте перейдите в **"APIs & Services"** → **"Library"**
2. Найдите и включите **Google Drive API**:
   - В поиске введите "Google Drive API"
   - Нажмите на результат и нажмите **"Enable"**
3. Найдите и включите **Google Sheets API**:
   - В поиске введите "Google Sheets API"
   - Нажмите на результат и нажмите **"Enable"**

#### Шаг 3: Создание учётных данных
1. Перейдите в **"APIs & Services"** → **"Credentials"**
2. Нажмите **"+ CREATE CREDENTIALS"** → **"OAuth client ID"**
3. Если появляется предупреждение о настройке consent screen:
   - Нажмите **"Configure Consent Screen"**
   - Выберите **"External"** (Внешний)
   - Заполните обязательные поля:
     - App name: "Excel to Sheets Converter"
     - User support email: ваш email
     - Developer contact: ваш email
   - Нажмите **"Save and Continue"**
   - Пропустите раздел "Scopes" (нажмите "Save and Continue")
   - Добавьте тестового пользователя (ваш email) в разделе "Test users"
   - Нажмите **"Save and Continue"**
4. Вернитесь к созданию OAuth Client ID:
   - Application type: **"Web application"**
   - Name: "Excel to Sheets Web Client"
   - В разделе **"Authorized JavaScript origins"** добавьте:
     - `http://localhost:5173` (для локальной разработки)
     - `https://your-domain.vercel.app` (для продакшена на Vercel)
   - В разделе **"Authorized redirect URIs"** добавьте:
     - `http://localhost:5173`
     - `https://your-domain.vercel.app`
   - Нажмите **"Create"**

#### Шаг 4: Получение ключей
После создания вы увидите окно с:
- **Your Client ID** - скопируйте это значение
- **Your Client Secret** (не требуется для этого приложения)

Также вам понадобится API Key:
1. В том же разделе "Credentials" нажмите **"+ CREATE CREDENTIALS"** → **"API key"**
2. Скопируйте полученный ключ
3. (Опционально) Нажмите **"Edit API key"** чтобы ограничить использование ключа только вашими доменами

#### Шаг 5: Настройка приложения
Откройте файл `src/App.jsx` и замените:
```javascript
const GOOGLE_API_KEY = "YOUR_API_KEY_HERE"
const GOOGLE_CLIENT_ID = "YOUR_CLIENT_ID_HERE"
```

На ваши значения:
```javascript
const GOOGLE_API_KEY = "AIzaSyD...ваш_ключ..."
const GOOGLE_CLIENT_ID = "123456789-abc...ваш_client_id...apps.googleusercontent.com"
```

### 3. Запуск приложения
```bash
npm run dev
```

Приложение будет доступно по адресу http://localhost:5173

### 4. Развёртывание на Vercel
```bash
# Установите Vercel CLI если ещё не установлен
npm i -g vercel

# Разверните проект
vercel
```

**Важно:** После развёртывания добавьте домен Vercel в разрешённые origin URI в Google Cloud Console!

## 📋 Как использовать

1. **Загрузите Excel файл** - перетащите файл .xlsx в зону загрузки или кликните для выбора
2. **Проверьте файл** - убедитесь, что все листы отображаются корректно
3. **Авторизуйтесь в Google** - нажмите кнопку входа и разрешите доступ к Google Диску
4. **Выберите папку** - выберите папку на Google Диске для сохранения таблицы
5. **Создайте таблицу** - нажмите кнопку создания и дождитесь завершения
6. **Откройте таблицу** - кликните по ссылке для открытия в Google Таблицах

## 🔧 Технологии

- **React** - фронтенд фреймворк
- **Vite** - сборщик проектов
- **xlsx** - библиотека для работы с Excel файлами
- **Google APIs** - Drive API и Sheets API
- **Tailwind CSS** - стилизация

## ⚠️ Важные замечания

1. **Безопасность**: Не коммитьте реальные API ключи в Git! Используйте переменные окружения для продакшена.

2. **Ограничения Google API**:
   - Бесплатный лимит: 100 запросов в 100 секунд для пользователя
   - Максимальный размер файла: зависит от квот вашего аккаунта

3. **Поддерживаемые форматы**: Только .xlsx (Excel 2007+)

## 📄 Лицензия

MIT
