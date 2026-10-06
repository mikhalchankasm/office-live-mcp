# Реестр → акты Word/PDF с ведущими нулями

**Запрос:** «Создай по тестовому реестру два акта и PDF в выбранной папке.»

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
      "Client"
    ],
    [
      "007",
      "Тестовый клиент А"
    ],
    [
      "0008",
      "Тестовый клиент Б"
    ]
  ],
  "value_mode": "text"
}
```

`excel_save_as` — обновить `${book}` по ответу

```json
{
  "workbook": "${book}",
  "path": "${folder}/register.xlsx"
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
  "text": "Акт {{Code}}\nКлиент: {{Client}}"
}
```

`word_save_as` — путь становится `${template}`

```json
{
  "document": "${doc}",
  "path": "${folder}/template.docx"
}
```

`bridge_excel_to_word_preview` — показать пользователю план; выполнение — после его согласия

```json
{
  "workbook": "${book}",
  "sheet": "Data",
  "cells": "A1:B3",
  "template": "${template}",
  "output_dir": "${folder}/acts",
  "filename_column": "Code",
  "export_pdf": true
}
```

`bridge_excel_to_word_documents` — повторно проверяет актуальные данные

```json
{
  "workbook": "${book}",
  "sheet": "Data",
  "cells": "A1:B3",
  "template": "${template}",
  "output_dir": "${folder}/acts",
  "filename_column": "Code",
  "export_pdf": true
}
```

**Что проверить:** `matched` содержит Code/Client, `ready=true`; sample и имена файлов сохраняют 007/0008. После preview папки acts ещё нет. После выполнения два DOCX и два PDF; при `ok=false` сверить полный `created_files`/`created_pdfs`, не запускать повторно вслепую.
