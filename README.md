# Табель 2026

Система обліку робочого часу для співробітників. Замінює Excel-табель.

## Як використовувати

1. Відкрийте [index.html](index.html) у браузері (або через GitHub Pages)
2. Виберіть своє ім'я зі списку та введіть пароль
3. Заповніть табель покроково (8 кроків)

## Структура проекту

| Файл | Призначення |
|------|-------------|
| `index.html` | Готова форма для співробітників |
| `template_v2.html` | Шаблон (редагувати тут) |
| `generate_html.py` | Генератор — читає `tabel_data.json` → пише `index.html` |
| `tabel_data.json` | Довідник: співробітники, об'єкти, розділи |
| `apps_script.gs` | Google Apps Script backend (не в репозиторії) |

## Після змін у шаблоні

```bash
python generate_html.py
```

## Google Apps Script

Backend для запису в Google Sheets та Telegram-сповіщень.  
Файл `apps_script.gs` не входить в репозиторій (містить токени).  
Розгортається окремо через Google Apps Script Editor.

## Технології

- Vanilla HTML / CSS / JavaScript (без залежностей)
- Google Apps Script (backend)
- Google Sheets (сховище даних)
