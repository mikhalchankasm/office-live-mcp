# Каталог инструментов

Сгенерировано командой `python -m office_live tools --markdown`. Описания — на английском: их читает агент.
Вид `write (readonly: list)` — многоактный инструмент: в режиме `OFFICE_LIVE_MODE=readonly` доступны только перечисленные действия чтения.

### excel_core (33)

| Tool | Kind | What it does |
|---|---|---|
| `excel_list_workbooks` | read | List every open Excel workbook across ALL running Excel instances: name, path, saved flag, sheets, visibility, and which one is active. Call this first. Returns an empty list (not an error) when Excel has no workbooks. |
| `excel_workbook_info` | read | Detailed overview of one workbook: every sheet with its used range, visibility, tables, charts, pivot tables, shapes, protection and filters, plus defined names and calculation mode. |
| `excel_open_workbook` | open | Open an existing workbook file in Excel (starts Excel if needed). Macros are disabled during opening. If the file is already open, returns it instead of opening twice. |
| `excel_new_workbook` | write | Create a new empty workbook in Excel (starts Excel if needed). The workbook is unsaved until excel_save_as. |
| `excel_close_workbook` | destructive | Close an open workbook. Refuses when it has unsaved changes unless you choose: save=true (saves first, needs an existing path) or discard=true (throws the changes away). Never quits Excel itself. |
| `excel_save` | save | Save a workbook to its current file (Ctrl+S). A workbook that was never saved needs excel_save_as instead. |
| `excel_save_as` | save | Save a workbook under a new name/format. Format follows the extension: .xlsx .xlsm .xlsb .xls .csv (UTF-8) .txt .html .ods. Refuses to replace an existing file unless overwrite=true. Afterwards the workbook is known by its NEW file name. |
| `excel_export_pdf` | save | Export a workbook (or one sheet) to a PDF file using its print settings. |
| `excel_read_range` | read | Read the values of a range. Dates come back as ISO strings, errors as '#DIV/0!' etc. Whole-column/row references and an empty `cells` are clipped to the used range. Large reads are truncated at max_cells and the response tells you the remaining range. |
| `excel_find` | read | Search cell values or formulas. Returns matching cells with their address and content. |
| `excel_get_selection` | read | What the user currently has selected in Excel (active workbook/sheet/range) and the values in it. Use this for requests like 'format what I selected' or 'sum this column'. |
| `excel_select_range` | ui | Activate a sheet and select (and scroll to) a range so the user sees it - useful to point at a result. |
| `excel_write_range` | write | Write a block of values. Pass a single top-left cell (e.g. 'B2') and the block is placed from there; rows shorter than the widest are padded with blanks; a single value fills the whole target range. Returns a read-back. |
| `excel_set_formula` | write | Put formulas into a cell or fill a whole range. A single formula string assigned to a multi-cell range is filled like Excel's fill-down: relative references adjust per cell (write '=A2*B2' with cells='C2:C100'). A 2-D array assigns each cell its own formula. Dynamic-array formulas (SORT, FILTER, UNIQUE, SEQUENCE) spill normally. Returns computed values and any error cells. |
| `excel_clear_range` | destructive | Clear parts of a range. |
| `excel_replace` | write | Replace text inside cell contents/formulas (keeps formatting). Returns how many cells and occurrences were affected. |
| `excel_add_worksheet` | write | Add a worksheet. Fails if the name is taken. By default the user's active sheet is left unchanged. |
| `excel_manage_sheet` | write | Rename, copy, move, hide/unhide, color or activate a sheet. |
| `excel_delete_sheet` | destructive | Permanently delete a sheet with all its data. Requires confirm=true. Refuses to delete the last sheet. |
| `excel_insert_rows_columns` | write | Insert empty rows or columns; existing content shifts down/right. |
| `excel_delete_rows_columns` | destructive | Delete rows or columns together with their content; the rest shifts up/left. References to deleted cells become #REF!. |
| `excel_set_dimensions` | write | Set column width / row height or autofit them (to hide rows/columns use excel_hide_rows_columns). |
| `excel_hide_rows_columns` | write | Hide or unhide whole rows/columns - either by number or by a condition on the data ('hide every row whose Status is Done'). Hidden rows keep their data; unhide with hidden=false. |
| `excel_group_rows_columns` | write | Create or remove outline groups (the +/- buttons in the margin) over rows or columns, or collapse/expand every group on the sheet. |
| `excel_copy_range` | write | Copy (or move) a range to another place, optionally into another sheet or workbook. |
| `excel_sort_range` | write | Sort a block of data by one or more columns (rows move as a whole). |
| `excel_filter_range` | write | Apply AutoFilter conditions (rows are hidden, not deleted) or clear them. |
| `excel_remove_duplicates` | destructive | Delete duplicate rows from a block of data, keeping the first occurrence. Comparison is case-insensitive like Excel's own command. |
| `excel_merge_cells` | write | Merge a range into one cell (only the top-left value survives) or unmerge it. |
| `excel_manage_names` | write (readonly: list) | List, add or delete defined names (named ranges / named formulas). |
| `excel_calculate` | write | Force recalculation and/or change the calculation mode (affects the whole Excel instance). |
| `excel_create_from_template` | save | Make a new workbook as an exact copy of a sample file (all formatting, merged cells, column widths, conditional formats, formulas, named ranges, and macros are kept) and open it. The sample itself is never modified. Then clear the old data and write the new data. For .xlsm samples keep the .xlsm extension; macros stay disabled in the copy until the user enables them in Excel. |
| `excel_autofill` | write | Do what dragging the fill handle does: extend a pattern from `source` over `dest` (which must contain `source`). Formulas adjust their relative references, numbers/dates continue as a series, formats are copied. Use it to carry a template row's formulas AND formatting down to new rows. |

### excel_format (10)

| Tool | Kind | What it does |
|---|---|---|
| `excel_format_range` | write | Format a range. Only the properties you pass are changed. Colors: '#RRGGBB' or a name (red, lightblue, lightgreen, lightyellow, ...); fill_color='none' removes the fill. |
| `excel_get_format` | read | Read the formatting of a range. A property is reported as 'mixed' when the cells of the range differ - narrow the range to see each value. |
| `excel_conditional_format` | write (readonly: list) | Add, list or clear conditional formatting ('highlight cells that ...'). |
| `excel_data_validation` | write (readonly: get) | Restrict what can be typed into cells (dropdown lists, number/date limits, custom rules), read the rule, or remove it. |
| `excel_add_hyperlink` | write | Turn a cell into a hyperlink - to a web/mail address or to another place in the workbook. |
| `excel_manage_comments` | write (readonly: list) | List, add/replace or delete cell notes (the yellow-sticker comments). |
| `excel_insert_image` | write | Insert a picture file (png/jpg/gif/bmp/emf) onto a sheet, anchored at a cell. |
| `excel_sheet_view` | write | Change how a sheet looks on screen: freeze panes, zoom, gridlines, row/column headings. |
| `excel_page_setup` | write | Print settings of a sheet (the needed printer driver must be available to Windows). |
| `excel_render_range_image` | read | Render a range exactly as it looks on screen (fonts, fills, borders, conditional formats) and return it as a PNG image - use it to visually verify formatting. Briefly uses the Windows clipboard. |

### word_core (20)

| Tool | Kind | What it does |
|---|---|---|
| `word_list_documents` | read | List every open Word document across ALL running Word instances: name, path, saved flag, page/paragraph counts, and which one is active. Call this first. Returns an empty list (not an error) when Word has no documents. |
| `word_get_structure` | read | Overview of a document: counts (pages, words, paragraphs, tables, images, comments, tracked changes, bookmarks, fields), the heading outline with paragraph numbers, a summary of every table, and sections. Use it to navigate a long document before reading parts of it. |
| `word_open_document` | open | Open an existing document file in Word (starts Word if needed). Macros are disabled while opening. If the file is already open it is returned instead of opened twice. |
| `word_new_document` | write | Create a new blank document in Word (starts Word if needed). The document is unsaved until word_save_as. |
| `word_close_document` | destructive | Close an open document. Refuses when it has unsaved changes unless you choose: save=true (saves first, needs an existing path) or discard=true. Never quits Word itself. |
| `word_save` | save | Save a document to its current file (Ctrl+S). A document that was never saved needs word_save_as instead. |
| `word_save_as` | save | Save a document under a new name/format. Format follows the extension: .docx .doc .docm .rtf .txt .html .odt (use word_export_pdf for PDF). Refuses to replace an existing file unless overwrite=true. Afterwards the document is known by its NEW file name. |
| `word_export_pdf` | save | Export a document to a PDF file (the open document keeps its name and format). |
| `word_read_document` | read | Read document content. mode='text' returns plain text in pages of max_chars (continue with `offset` = previous offset + max_chars; the response says when the end is reached). mode='paragraphs' returns numbered paragraphs with their style, heading level and whether they sit in a table - use it to learn paragraph numbers for editing. |
| `word_find` | read | Find every occurrence of a text and report where it is (paragraph number, position) with surrounding context. |
| `word_fill_placeholders` | write | Fill a template: replace every {{key}} (also {{ key }}) with its value, in the body, headers, footers and text boxes. Reports placeholders left unfilled. |
| `word_replace_text` | write | Replace every occurrence of a text (formatting of the replaced text is kept). Literal by default: the characters '^' and '\' are NOT special. Returns the exact number of replacements. With 'Track changes' on, replacements are recorded as revisions. |
| `word_insert_text` | write | Insert text. By default each insertion becomes its own paragraph(s) with the given style (newlines in `text` make several paragraphs). Returns the paragraph numbers of the inserted text. |
| `word_delete_paragraphs` | destructive | Delete one paragraph or a run of paragraphs (with their text). |
| `word_format_text` | write | Apply a style and/or direct formatting. Pick the target with `scope`: paragraphs (start..end paragraph numbers), find (every occurrence of find_text), selection (what the user selected), or document (everything). Only the properties you pass are changed. Colors: '#RRGGBB' or names. |
| `word_list_styles` | read | List the styles available in a document (names are what word_format_text/word_insert_text accept). |
| `word_make_list` | write | Turn paragraphs into a bulleted or numbered list, or remove list formatting. |
| `word_get_selection` | read | What the user currently has selected (or where the cursor is) in a Word document: text, paragraph numbers, style, whether it is inside a table. Use it for requests like 'rewrite the paragraph I selected'. |
| `word_select` | ui | Move the user's cursor/selection and scroll there, to point at a result. |
| `word_create_from_template` | save | Make a new Word document as an exact copy of a sample (.docx/.docm/.dotx/.dotm: styles, headers, footers, tables, placeholders, macros kept) and open it. The sample is never modified. Then use word_fill_placeholders / word_insert_text / word_write_table to put the new content in. |

### word_layout (13)

| Tool | Kind | What it does |
|---|---|---|
| `word_page_setup` | write | Page layout: orientation, paper size, margins, text columns. |
| `word_headers_footers` | write (readonly: get) | Read or set page headers/footers: text and automatic page numbers. |
| `word_manage_toc` | write (readonly: list) | Insert, refresh or remove a table of contents built from the Heading styles. |
| `word_update_fields` | write | Refresh every field of the document, including headers/footers, footnotes and text boxes: page numbers, total pages, table of contents, cross-references, dates. Reports fields that could not be updated. |
| `word_insert_image` | write | Insert a picture file (png/jpg/gif/bmp/emf) as its own paragraph. |
| `word_insert_break` | write | Insert a page, section or column break. |
| `word_manage_comments` | write (readonly: list) | Review comments: list, add (on a paragraph or on found text), reply, resolve, delete. |
| `word_track_changes` | write (readonly: status, list) | Control 'Track Changes' and review tracked revisions. |
| `word_manage_bookmarks` | write (readonly: list) | List, add or delete bookmarks (named places you can later fill with word_insert_text(position='bookmark')). |
| `word_insert_hyperlink` | write | Make text a hyperlink: either link existing text (find_text) or append new linked `text` to the end of a paragraph. |
| `word_manage_footnotes` | write (readonly: list) | List, add or delete footnotes. |
| `word_document_properties` | write (readonly: get) | Read or set built-in document properties (File > Info): Title, Subject, Author, Keywords, Comments, Category, Company, Manager. |
| `word_render_page_image` | read | Render one page of a Word document exactly as laid out (fonts, tables, pictures, headers/footers) and return it as a PNG image - use it to visually verify a document. Needs the Pillow package. The user's document window is switched to Print Layout briefly and restored. |

### word_tables (6)

| Tool | Kind | What it does |
|---|---|---|
| `word_list_tables` | read | List the tables of a document: index, size, style, first cell, and the paragraph number where each table starts. |
| `word_read_table` | read | Read a table as a grid of cell texts (rows x columns). Tables with merged cells come back as a list of {row, col, text} cells instead. |
| `word_create_table` | write | Create a table from a 2-D array of values (the first row is the header by default) and return its index. |
| `word_write_table` | write | Write a block of values into an existing table starting at (start_row, start_col). Rows are appended when the data runs past the last row (add_rows=true). Not supported for tables with merged cells in the target area. |
| `word_modify_table` | write | Change a table's structure. |
| `word_format_table` | write | Format a table or a block of its cells. Only the properties you pass are changed. Cell block = row_from..row_to x col_from..col_to (all zero = the whole table). |

### bridge (7)

| Tool | Kind | What it does |
|---|---|---|
| `bridge_word_table_to_excel` | write | Copy a table from a Word document into an Excel sheet. Numbers typed as text in Word ('1 234,50', '12%') become real Excel numbers, everything else stays text (IDs like 007 keep their zeros). Works across Word/Excel instances. |
| `bridge_word_text_to_excel` | write | List the paragraphs of a Word document in an Excel sheet as rows [paragraph no, style, text] - handy for analysing, tagging or restructuring a document's text. |
| `bridge_excel_range_to_word_table` | write | Copy an Excel range into a Word document as a real, editable Word table. By default the text is exactly what Excel shows (number formats, dates, percentages), hidden rows/columns are skipped, and numeric columns are right-aligned. |
| `bridge_excel_range_to_word_picture` | write | Paste an Excel range into Word as a PICTURE that looks exactly as on the Excel screen (fills, borders, conditional formats, fonts). Not editable text - use bridge_excel_range_to_word_table for that. Briefly uses the Windows clipboard. |
| `bridge_excel_chart_to_word` | write | Insert an Excel chart into a Word document as a picture. |
| `bridge_excel_to_word_documents` | save | Generate one Word document per Excel row from a Word template containing {{Header}} placeholders (the header names are the first row of `cells`). Typical use: letters, contracts, certificates, acts. Each document is saved into output_dir (and optionally as PDF); the template is never changed. |
| `office_inspect_file` | read | Analyse an .xlsx/.xlsm/.docx/.docm file WITHOUT opening it in Office: sheets and their size, merged ranges, frozen panes, hidden columns, conditional-format rules, validations, tables/charts/pivots, defined names, the first rows of content, and the macros (module and procedure names, optionally the VBA source code - read only, never executed). Use it first to understand a sample file. |

### eval (1)

| Tool | Kind | What it does |
|---|---|---|
| `office_run_python` | destructive | ADVANCED / DANGEROUS: run Python code with live COM objects when no dedicated tool covers a need. Predefined names: `excel` (Excel.Application or None), `wb` (workbook or None), `ws` (its active sheet), `word` (Word.Application or None), `doc` (document or None), `pythoncom`, `win32com`. Set `result = ...` to return a value; print() output is returned too. COM rules: pass arguments POSITIONALLY (no name=value), use None for skipped optional arguments, and avoid Range.Resize/Offset (they misbehave in late binding). Only available when the server was started with OFFICE_LIVE_ALLOW_EVAL=1. |

### excel_analysis (13)

| Tool | Kind | What it does |
|---|---|---|
| `excel_manage_tables` | write (readonly: list) | Work with Excel tables (the structured 'Format as Table' objects with filter buttons and auto-growing ranges). |
| `excel_create_pivot_table` | write | Create a PivotTable ('сводная таблица') from a data block or an Excel table. The first row of the source must contain unique header names. |
| `excel_manage_pivot_tables` | write (readonly: list) | List, refresh or delete pivot tables. |
| `excel_create_chart` | write | Create a chart from a data block. Include the header row and the label column in `source`: the first column becomes the category axis and each other column a series (use series_in_rows=true if the series are laid out in rows). |
| `excel_manage_charts` | write (readonly: list, export_image) | List, restyle, move, delete or export charts. |
| `excel_profile_range` | read | Statistical profile of a data block, column by column: types, empty/unique counts, min/max/mean/median/sum for numbers, date range, most frequent values, duplicates, formula counts. Use it to understand unfamiliar data before analysing or cleaning it. |
| `excel_find_issues` | read | Audit a sheet/range for problems: error values, numbers stored as text, inconsistent formulas down a column, hard-coded numbers among formulas, stray spaces, blank rows/headers, duplicate headers, mixed types in a column, merged cells. |
| `excel_pivot_info` | read | Describe a pivot table in detail: every field with its role (row/column/filter/data/unused), the items of each field and whether they are visible (hidden = filtered out), data fields with their function, and applied filters. Use it before filtering or restructuring. |
| `excel_pivot_filter` | write (readonly: items) | Filter a pivot table - hide/show rows (items) of a field, apply label/value/top-N filters, or clear filters. This is what the field's filter dropdown does in Excel. |
| `excel_pivot_fields` | write | Restructure a pivot table: add/move/remove fields, change how a value is aggregated, sort, group dates or numbers, add calculated fields, expand/collapse. |
| `excel_pivot_options` | write | Layout and behaviour options of a pivot table. |
| `excel_manage_slicers` | write (readonly: list) | Create and drive slicers (clickable filter buttons, 'срезы') and date timelines for pivot tables and Excel tables. |
| `excel_describe_layout` | read | Produce a BLUEPRINT of a sheet so it can be re-created with new data: where the title/header block ends (frozen rows), merged header cells, every column (header text, width, hidden, number format, data type, typical values, the formula it contains), the KINDS of data rows (e.g. section rows vs item rows - grouped by how they actually look on screen incl. conditional formatting - with examples and code patterns), conditional-format rules in English with their colors, hidden rows, tables/charts/pivots/validation/macros. Read it, then build the analogue. |
