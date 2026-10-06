<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/assets/banner-ru-dark.svg">
    <img alt="Office Live MCP — ваш ИИ-агент в уже открытых файлах Excel и Word" src="docs/assets/banner-ru-light.svg" width="100%">
  </picture>
</p>

<p align="center">
  <a href="https://github.com/mikhalchankasm/office-live-mcp/releases/latest"><img alt="Последний релиз" src="https://img.shields.io/github/v/release/mikhalchankasm/office-live-mcp?label=%D1%80%D0%B5%D0%BB%D0%B8%D0%B7&color=217346"></a>
  <a href="https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml"><img alt="CI" src="https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg"></a>
  <a href="LICENSE"><img alt="Лицензия MIT" src="https://img.shields.io/badge/license-MIT-2b579a"></a>
  <img alt="Windows 10/11" src="https://img.shields.io/badge/Windows-10%20%7C%2011-555">
</p>

<p align="center">
  <a href="https://github.com/mikhalchankasm/office-live-mcp/releases/latest"><img alt="Скачать установщик Windows" src="https://img.shields.io/badge/%D0%A1%D0%BA%D0%B0%D1%87%D0%B0%D1%82%D1%8C_%D0%B4%D0%BB%D1%8F_Windows-setup.exe-217346?style=for-the-badge"></a>
  &nbsp;
  <a href="docs/DETAILS.ru.md"><img alt="Документация" src="https://img.shields.io/badge/%D0%94%D0%BE%D0%BA%D1%83%D0%BC%D0%B5%D0%BD%D1%82%D0%B0%D1%86%D0%B8%D1%8F-%E2%86%92-2b579a?style=for-the-badge"></a>
</p>

<p align="center"><a href="README.md">English</a> · <b>Русский</b></p>

<h3 align="center">Ваш ИИ-агент. Ваши открытые Excel и Word. Меньше ручной работы.</h3>

<p align="center">
Превращайте строки реестра в документы Word/PDF, сравнивайте таблицы и согласовывайте правки Word —<br>
через привычного агента, в уже открытых файлах.
</p>

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<p align="center"><sub>30 секунд: настоящие вызовы инструментов в настоящих Excel и Word. Данные вымышленные.</sub></p>

## Почему Office Live MCP

<table>
  <tr>
    <td width="33%" valign="top">
      <h4>🤖&nbsp; Свой агент и модель</h4>
      Сохраняйте свои инструкции и привычный порядок работы. Любая модель, которую поддерживает ваш MCP-клиент.
    </td>
    <td width="33%" valign="top">
      <h4>📄&nbsp; Excel → Word → PDF</h4>
      Документ на каждую строку реестра по вашему шаблону — с предпросмотром имён файлов и конфликтов.
    </td>
    <td width="33%" valign="top">
      <h4>🛡️&nbsp; Доступ под контролем</h4>
      Режим только чтения, разрешённые папки и набор инструментов — их проверяет сервер, а не подсказка модели.
    </td>
  </tr>
  <tr>
    <td width="33%" valign="top">
      <h4>👀&nbsp; Правки сразу видны</h4>
      Работа в уже открытых документах, включая несохранённые правки. Результат сразу на экране.
    </td>
    <td width="33%" valign="top">
      <h4>↩️&nbsp; Просмотр и отмена</h4>
      Правки Word — с предпросмотром и через исправления. Журнал и отмена с проверкой конфликтов.
    </td>
    <td width="33%" valign="top">
      <h4>🔓&nbsp; Бесплатно и открыто</h4>
      Без подписки на сервер, код под MIT. Один установщик — без Python и прав администратора.
    </td>
  </tr>
</table>

<p align="center">
  <b>Работает с</b>&nbsp; Claude Code · Claude Desktop · Cursor · Codex CLI · ZCode · VS Code<br>
  <sub>и другими MCP-клиентами</sub>
</p>

## Попробуйте на конкретной задаче

| Попросите агента | Получите |
|---|---|
| *«Создай документы Word и PDF по этому реестру и моему шаблону. Сначала покажи план».* | [Документ на каждую строку с предпросмотром до генерации](docs/examples/register-to-acts.md). Коды вроде `007` остаются текстом. |
| *«Сравни эти два реестра по коду и покажи, что изменилось».* | [Расхождения с сопоставлением по ключу](docs/examples/compare-registers.md). |
| *«Прочитай замечания в Word и предложи правки. Внеси после моего согласования».* | [Точечные правки, записанные через исправления](docs/examples/comments-to-tracked-edits.md). |

**123 инструмента** для Excel, Word и работы с документами — формулы, очистка данных, подбор параметра, сводные,
диаграммы, шаблоны, ветки примечаний и другое. [Весь каталог →](docs/TOOLS.md)

## Установка

> **Что нужно:** Windows 10/11 (64-бит) и настольные Excel и/или Word.

1. **Скачайте** установщик (`…-setup.exe`) со страницы [последнего релиза](https://github.com/mikhalchankasm/office-live-mcp/releases/latest).
2. **Запустите:** Далее → Далее → Готово. При первой установке найденные агенты уже отмечены; выберите доступ: **полный** или **только чтение**.
3. **Перезапустите агента** и попросите: *«покажи, какие книги открыты в Excel»*.

<sub>**Обновление:** новый setup.exe поверх старого. **Удаление:** Параметры → Приложения → Office Live MCP.
**Данные и стоимость:** сервер работает локально, без надстройки Office; прочитанное агентом получает поставщик его модели,
агент и модель оплачиваются по условиям их поставщика.</sub>

## Подробнее

[Подробная документация](docs/DETAILS.ru.md) · [Каталог инструментов](docs/TOOLS.md) · [Руководство](docs/GUIDE.ru.md) ·
[Изменения](CHANGELOG.md) · [Сравнение с Claude, ChatGPT и другими MCP-серверами](docs/DETAILS.ru.md#сравнение-с-аналогами) ·
[Лицензия MIT](LICENSE)
