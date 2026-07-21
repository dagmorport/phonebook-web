# Телефонный справочник «Зеленое яблоко»

Корпоративный веб- справочник сотрудников компании Pepper Group.

## Возможности

- Поиск по имени, фамилии, должности, отделу
- Алфавитный указатель для быстрой навигации
- Тёмная тема
- Избранное (сохраняется в браузере)
- Копирование email одним кликом
- Автоматическая синхронизация с Active Directory

## Технологии

- **HTML5** — структура приложения
- **Tailwind CSS** — утилитарные стили, адаптивная вёрстка
- **Lucide Icons** — SVG-иконки
- **Vanilla JavaScript** — логика приложения
- **PowerShell 5.1** — скрипт экспорта из Active Directory

## Установка

1. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/your-username/phonebook-web.git
   ```

2. Откройте `index.html` в браузере

## Экспорт из Active Directory

Для обновления данных сотрудников запустите:

```powershell
.\Export-Phonebook.ps1
```

Скрипт подключается к Active Directory и обновляет файл `employees.js`.

## Структура проекта

```
phonebook-web/
├── index.html              # Основное приложение
├── employees.js            # Данные сотрудников (из AD)
├── dept-filter.json        # Фильтр отделов
├── Export-Phonebook.ps1    # Скрипт экспорта
├── Export-Phonebook.bat    # Обёртка для запуска
├── Phonebook.ps1           # WPF-приложение (десктоп)
├── Logo.png                # Логотип
└── CHANGELOG.md            # История изменений
```

## Лицензия

MIT
