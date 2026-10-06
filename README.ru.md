# Office Live MCP

[![CI](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg)](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.10%E2%80%933.13-blue)

[![English](https://img.shields.io/badge/English-README-0969da?style=for-the-badge)](README.md)
[![Русский](https://img.shields.io/badge/%D0%A0%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9-README-cf222e?style=for-the-badge)](README.ru.md)

MCP-сервер с открытым кодом (MIT), через который **любой MCP-агент** — Claude Code, Claude Desktop, Cursor, Codex CLI,
ZCode, VS Code… — работает в документах Excel и Word, **которые у вас уже открыты**. Изменения сразу видны на экране, а
вы продолжаете работать рядом с агентом. Подключается к запущенным Excel и Word через COM — без надстройки Office.

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<sub>30 секунд: настоящие вызовы инструментов в настоящих Excel и Word. Данные вымышленные.</sub>

## Главное

- **Работает в уже открытых файлах** — видит несохранённые правки, пересчитанные формулы, сводные, диаграммы и текущее выделение.
- **Excel и Word, 123 инструмента** — диапазоны, формулы, чистка данных, подбор параметра, промежуточные итоги, сводные со срезами, диаграммы; структура Word, шаблоны, правки исправлениями, ветки примечаний, сравнение документов.
- **Мосты** — таблицы Word → Excel с настоящими числами, диапазоны и диаграммы Excel → Word, документ Word/PDF на каждую строку Excel — с предварительным просмотром без изменений.
- **Ограничения задаёте вы** — режим «только чтение», разрешённые папки, наборы инструментов и строгий выбор цели; ничего не сохраняется само; локальный журнал и отмена поддерживаемых действий, которая отказывает, если вы потом правили то же место.
- **Любой MCP-агент, один установщик** — без Python, git и прав администратора. Сервер бесплатный и работает локально; агент и его модель оплачиваются по условиям их поставщика, и прочитанное агентом уходит к нему.

## Установка

**Что нужно:** Windows 10/11 (64-бит) и настольные Excel и/или Word.

1. Скачайте **`office-live-mcp-<версия>-setup.exe`** со страницы [**Releases**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest).
2. Запустите: **Далее → Далее → Готово** — выберите агентов и доступ: **полный** или **только чтение**.
3. **Перезапустите агента** и попросите: *«покажи, какие книги открыты в Excel»*.

**Обновление:** запустите новый setup.exe поверх установленного и перезапустите агента.
**Удаление:** **Параметры → Приложения → Office Live MCP → Удалить**.

## Подробнее

[![Подробная документация](https://img.shields.io/badge/%D0%9F%D0%BE%D0%B4%D1%80%D0%BE%D0%B1%D0%BD%D0%B0%D1%8F_%D0%B4%D0%BE%D0%BA%D1%83%D0%BC%D0%B5%D0%BD%D1%82%D0%B0%D1%86%D0%B8%D1%8F-%E2%86%92-cf222e?style=for-the-badge)](docs/DETAILS.ru.md)

Тихая установка и параметры setup, запасной ZIP, все возможности, окна и кликабельные ссылки, безопасность, сравнение
с ИИ-надстройками Office и другими MCP-серверами, разработка — [docs/DETAILS.ru.md](docs/DETAILS.ru.md).
Все 123 инструмента: [docs/TOOLS.md](docs/TOOLS.md) · Руководство: [docs/GUIDE.ru.md](docs/GUIDE.ru.md) ·
Изменения: [CHANGELOG.md](CHANGELOG.md) · Лицензия: [MIT](LICENSE)
