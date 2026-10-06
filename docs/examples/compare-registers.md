# Две версии реестра → таблица различий

**Запрос:** «Сравни два синтетических реестра по коду и запиши различия в новую книгу.»

Обозначения: `${folder}` — папка, которую выбрал пользователь и разрешил серверу; `${book}`, `${doc}`, `${template}` — точные имена/пути из ответов предыдущих вызовов. Подставьте значения перед вызовом. Данные полностью синтетические; имена результатов в примере не передаются буквально. Перед первой записью в книгу предложите журнал `Лог` согласно INSTRUCTIONS.

**Последовательность:**

`excel_new_workbook` — сохранить workbook как `${book}`

```json
{
  "sheets": [
    "Before",
    "After"
  ]
}
```

`excel_write_range`

```json
{
  "workbook": "${book}",
  "sheet": "Before",
  "cells": "A1",
  "values": [
    [
      "Code",
      "Status"
    ],
    [
      "007",
      "open"
    ],
    [
      "008",
      "open"
    ]
  ],
  "value_mode": "text"
}
```

`excel_write_range`

```json
{
  "workbook": "${book}",
  "sheet": "After",
  "cells": "A1",
  "values": [
    [
      "Code",
      "Status"
    ],
    [
      "008",
      "closed"
    ],
    [
      "007",
      "open"
    ]
  ],
  "value_mode": "text"
}
```

`excel_save_as` — обновить `${book}`

```json
{
  "workbook": "${book}",
  "path": "${folder}/versions.xlsx"
}
```

`excel_compare_ranges` — получить различия; убедиться, что ответ не обрезан

```json
{
  "workbook_a": "${book}",
  "sheet_a": "Before",
  "cells_a": "A1:B3",
  "workbook_b": "${book}",
  "sheet_b": "After",
  "cells_b": "A1:B3",
  "match": "key",
  "key_column": "Code"
}
```

`excel_new_workbook` — сохранить workbook как `${result_book}`

```json
{
  "sheets": [
    "Differences"
  ]
}
```

`excel_write_range` — значения здесь воспроизводят единственное различие синтетических данных из ответа

```json
{
  "workbook": "${result_book}",
  "sheet": "Differences",
  "cells": "A1",
  "values": [
    [
      "Code",
      "Column",
      "Before",
      "After"
    ],
    [
      "008",
      "Status",
      "open",
      "closed"
    ]
  ],
  "value_mode": "text"
}
```

`excel_save_as`

```json
{
  "workbook": "${result_book}",
  "path": "${folder}/differences.xlsx"
}
```

**Что проверить:** Перестановка строк сама по себе не считается отличием. Единственное различие — статус 008; ведущие нули остались текстом, строки сопоставлены по ключу. В реальных вызовах таблицу строить из differences ответа, а при truncated дочитать/сузить диапазон.
