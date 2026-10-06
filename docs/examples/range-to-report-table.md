# Диапазон Excel → редактируемая таблица Word

**Запрос:** «Сделай тестовый отчёт с редактируемой таблицей из Excel.»

Обозначения: `${folder}` — папка, которую выбрал пользователь и разрешил серверу; `${book}`, `${doc}`, `${template}` — точные имена/пути из ответов предыдущих вызовов. Подставьте значения перед вызовом. Данные полностью синтетические; имена результатов в примере не передаются буквально. Перед первой записью в книгу предложите журнал `Лог` согласно INSTRUCTIONS.

**Последовательность:**

`excel_new_workbook` — сохранить workbook как `${book}`

```json
{
  "sheets": [
    "Data"
  ]
}
```

`excel_write_range`

```json
{
  "workbook": "${book}",
  "sheet": "Data",
  "cells": "A1",
  "values": [
    [
      "Code",
      "Quantity"
    ],
    [
      "007",
      10
    ],
    [
      "0008",
      20
    ]
  ],
  "value_mode": "text"
}
```

`excel_save_as` — обновить `${book}`

```json
{
  "workbook": "${book}",
  "path": "${folder}/report-data.xlsx"
}
```

`word_new_document` — сохранить document как `${doc}`

```json
{}
```

`word_insert_text`

```json
{
  "document": "${doc}",
  "text": "Синтетический отчёт",
  "style": "Heading 1"
}
```

`bridge_excel_range_to_word_table`

```json
{
  "workbook": "${book}",
  "sheet": "Data",
  "cells": "A1:B3",
  "document": "${doc}",
  "position": "end",
  "header_row": true,
  "keep_formatting": true
}
```

`word_list_tables` — проверить номер таблицы из ответа моста

```json
{
  "document": "${doc}"
}
```

`word_save_as`

```json
{
  "document": "${doc}",
  "path": "${folder}/report.docx"
}
```

**Что проверить:** Таблица 3×2 содержит 007 и 0008, числа выровнены как числа, первая строка — заголовок; текст в ячейке можно редактировать. В ответе моста нет пропущенных скрытых строк/столбцов. Это таблица Word, не изображение; office_undo отменяет вставку до отдельного сохранения.
