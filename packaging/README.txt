Office Live MCP — управление открытыми Excel и Word из ИИ-агентов (Claude, Cursor, Codex, ZCode…)

УСТАНОВКА
1. Распакуйте архив целиком (правой кнопкой — «Извлечь все…»). Запускать прямо из архива нельзя.
2. Дважды щёлкните install.cmd.
3. Установщик по шагам проверит Windows и Office, скопирует программу в
   %LOCALAPPDATA%\Programs\office-live-mcp (права администратора не нужны), проверит её и спросит:
   — какой доступ дать агентам: полный или только чтение;
   — к каким агентам подключить (найденные на компьютере отмечены).
4. Перезапустите агента и попросите, например: «покажи, какие книги открыты в Excel».

Что нужно: Windows 10/11 (64-бит) и настольные Excel и/или Word. Python, winget и прочее не нужны — всё внутри.
Сам агент работает через интернет, программа Office Live — только на этом компьютере.

ОБНОВЛЕНИЕ
Распакуйте новый архив и снова запустите install.cmd. Каталог и регистрации запоминаются в
%LOCALAPPDATA%\office-live-mcp\state.json. Настройки сохраняются, новые агенты не подключаются автоматически.
Для старой нестандартной установки без state при первом обновлении укажите --target "D:\Office Live".
Если установщик скажет, что программа занята, закройте агентов (Claude, Cursor…) и повторите.

БЕЗ ВОПРОСОВ (для администраторов)
install.cmd --yes --clients claude-code,cursor --readonly

УДАЛЕНИЕ
"%LOCALAPPDATA%\Programs\office-live-mcp\app\office-live-mcp.exe" uninstall

ЕСЛИ WINDOWS ИЛИ АНТИВИРУС БЛОКИРУЕТ ЗАПУСК
Программа не подписана сертификатом. SmartScreen («Windows защитила ваш компьютер») — «Подробнее» → «Выполнить в любом случае».
Если включён Smart App Control или корпоративный антивирус, попросите администратора разрешить папку программы.


ТОЛЬКО ПРОГРАММА: install.cmd --yes --clients none
ПЛАН БЕЗ ИЗМЕНЕНИЙ: install.cmd --dry-run --yes --clients none
ПОЗДНЕЕ ПОДКЛЮЧЕНИЕ: app\office-live-mcp.exe setup --yes --clients codex,zcode
РАЗНЫЙ ДОСТУП: setup --yes --clients codex --full, затем setup --yes --clients zcode --readonly
ПРОЕКТ: setup --clients cursor --scope project --project "D:\Проект"
СВОЙ КОНФИГ: setup --clients zcode --config "D:\Конфиги\zcode.json"
ОТКЛЮЧИТЬ ОДНОГО: setup --yes --remove --clients zcode
Установленный EXE: %LOCALAPPDATA%\Programs\office-live-mcp\app\office-live-mcp.exe (или выбранный --target).
Конфигурация проверена чтением обратно; нужен перезапуск клиента. Загрузка им сервера не проверяется автоматически.
При удалении изменённые записи/файлы сохраняются. Если не удалось снять зависимую запись, EXE остаётся.
Итог удаления из самого EXE: %LOCALAPPDATA%\office-live-mcp\uninstall-report.json.
Portable и установка для всех пользователей не поддерживаются.
