---
name: xlsx
description: >-
  Use this skill any time a spreadsheet file is the primary input or output.
  Trigger when the user wants to open, read, edit, fix, or create .xlsx, .xlsm,
  .xltx, .csv, or .tsv; clean messy tabular data; convert between tabular
  formats; or produce a polished multi-sheet workbook (Resumo + detail tabs,
  reports, exports). Trigger especially when they name a spreadsheet path or ask
  for a planilha/relatório em Excel. The deliverable must be a spreadsheet file.
  Do NOT trigger when the primary deliverable is a Word document, HTML report,
  standalone Python script, database pipeline, PDF, or Google Sheets API
  integration, even if tabular data is involved.
license: Complete terms in LICENSE.txt
---

# XLSX creation, editing, and analysis

| Task | Approach |
|---|---|
| **Create** or **edit** with formulas/formatting | `openpyxl` — see gotchas below |
| **Bulk data** in or out | `pandas` (`read_excel`, `to_excel`) |
| **Quick look** at a sheet | `markitdown file.xlsx` — `## SheetName` per sheet; reads `.xlsm` too. No cell coordinates, so don't plan edits from it |
| **Read** a model (formulas *and* values) | two `load_workbook` passes — see gotchas |

> `openpyxl`, `pandas`, and `markitdown` are preinstalled — do not run `pip install` first; write the script and import directly. Only if an import fails (or the `markitdown` command is missing): `pip install` the missing package.

> Script paths below are relative to this skill's directory.

## Requirements for every output

- **Professional font** (Calibri or Arial) throughout, unless the user says otherwise.
- **Zero formula errors.** Never ship while `recalc.py` reports `errors_found`; the one exception is under `not_computed` below. If you think an error predates you, prove it: load the *original* with `data_only=True` and look at that cell. An error you introduced looks exactly like one you inherited.
- **Use formulas, never hardcoded results.** Write `sheet['B10'] = '=SUM(B2:B9)'`, not the Python-computed total. The sheet must recalculate when its inputs change.
- **Follow the user's spec literally.** Exact tab names, exact column headers, and the formula they spelled out. A redesign that computes something else fails, however elegant.
- **Document every assumption and hardcoded number** where the reader will see it — a cell comment, or an adjacent cell at a table's end. Cite a real source when one exists; when the number came from the user, say so plainly.
- **A workbook *you create* for someone to fill in** needs a short legend naming which cells to edit, and one example row of realistic values showing the expected format. Never add such a row to a file you were asked to edit.
- **Editing an existing file: match its conventions exactly.** They override every guideline here. Find its designated input cells first — a distinct font color, fill, or shading marks them — write only there, and leave every existing formula untouched.

## Delivery workbooks (new files for humans)

When **creating** a workbook as a deliverable (report, analysis, export, planilha de entrega), decide the sheet map **before** writing cells. Do not dump everything into one sheet and call it done.

### Sheet map

1. **Resumo** (first tab) — create when there are **2+ detail sheets**, or when the reader needs KPIs/totals before the raw rows. Skip Resumo for a single simple table with no rollups.
2. **Detail sheets** — one sheet per logical cut (period, channel, category, entity). Prefer several focused sheets over one mega-grid.
3. **Supporting sheets** (only if needed): `Premissas`, `Legenda`, or `Dados` for inputs, assumptions, or staging data the formulas depend on.
4. **Never** create empty decorative sheets. Sheet names ≤ 31 chars; never use `\ / ? * [ ]`.

Tab order: `Resumo` → detail sheets (most important first) → supporting sheets last.

### When Resumo is required

| Situation | Resumo? |
|---|---|
| One flat table, no totals story | No — single data sheet is enough |
| Multiple cuts the user will compare | Yes |
| KPIs, compliance rates, totals, rankings | Yes |
| Template for someone else to fill | Optional short Resumo + clear Legenda |

### What goes on Resumo

- Title block: workbook title, period/scope, generated date
- KPI block or small metrics table — **formulas** that reference detail sheets (`=SUM('Vendas'!E:E)`, `=COUNTA(...)`)
- At most **one** chart unless the user asks for more
- Short labels pointing to detail tabs (“Detalhe por canal → aba Canais”)
- **Never** paste the full raw dataset on Resumo

### Visual standard (delivery files)

Aim for “opened in a meeting and readable,” not dashboard eye-candy.

- **Type:** Calibri/Arial 11 body; 14 bold for sheet title; headers 11 bold
- **Header row:** bold, freeze panes below it, `AutoFilter` on
- **Zebra (optional):** light fill `#F2F2F2` on data rows — pick one style and keep it consistent across sheets
- **Header fill:** solid `#1F4E79` with white text, **or** light `#D9E2F3` with dark text — one choice per workbook
- **Borders:** thin hairline on data tables only; no thick colored boxes
- **Numbers:** dates `DD/MM/YYYY` (or ISO if the user prefers); currency with thousands separator; percentages stored as fractions with `0.0%`
- **Columns:** size to content, then clamp roughly 10–45 width; never leave headers truncated
- **Alignment:** text left, numbers/dates right, headers center or left consistently
- **Avoid:** emoji in headers, purple gradients, stacked decorative merges, giant logos, more than two accent colors

### Detail sheet layout

```
Row 1 (optional): sheet title spanning columns
Row 2 / Row 1:    column headers
Row 3+:           data
Last row:         totals with formulas when sums/counts matter
```

- Sort rows in a meaningful order (date, rank, name)
- Totals via formulas (`SUM`, `COUNT`, `AVERAGE`) — never paste Python aggregates as literals when source cells exist
- Cross-sheet references must quote names with spaces: `='Minha Aba'!B2`

### Minimal openpyxl pattern

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
BODY_FONT = Font(name="Calibri", size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
ZEBRA = PatternFill("solid", fgColor="F2F2F2")

def style_header(ws, row, start_col, end_col):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row + 1, 1).coordinate
    ws.auto_filter.ref = ws.dimensions

def autosize(ws, min_w=10, max_w=45):
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        length = max(len(str(c.value or "")) for c in col_cells)
        ws.column_dimensions[letter].width = max(min_w, min(max_w, length + 2))

wb = Workbook()
# Resumo first
resumo = wb.active
resumo.title = "Resumo"
resumo["A1"] = "Título do relatório"
resumo["A1"].font = TITLE_FONT
resumo["A2"] = f"Gerado em {date.today().isoformat()}"
resumo["A2"].font = BODY_FONT
# KPIs with formulas after detail sheets exist, e.g.:
# resumo["A4"] = "Total"
# resumo["B4"] = "=SUM('Detalhe'!E:E)"

detalhe = wb.create_sheet("Detalhe")
headers = ["Data", "Canal", "Valor"]
for i, h in enumerate(headers, 1):
    detalhe.cell(1, i, h)
style_header(detalhe, 1, 1, len(headers))
# ... write rows, then autosize(detalhe)
```

### Delivery checklist (before handing the file over)

- [ ] Sheet map matches the story (Resumo only if needed; no empty tabs)
- [ ] Resumo KPIs use formulas into detail sheets
- [ ] Freeze + filter on every data sheet
- [ ] Number formats applied (date / currency / %)
- [ ] Columns readable; headers not cut off
- [ ] `recalc.py` clean when formulas exist
- [ ] Opened once mentally as a first-time reader: first tab answers “so what?”

## Recalculate (mandatory whenever the file contains formulas)

openpyxl writes formulas as strings with **no cached values**. Until you recalculate, every
formula cell reads back as `None` to anything reading cached values — `pandas`,
`load_workbook(data_only=True)`, and most previewers.

```bash
python scripts/recalc.py output.xlsx [timeout_seconds]   # default 30
```

LibreOffice computes the formulas **on a copy**, only their values go into your file, and you get JSON:
`status` (`success` | `errors_found`), `total_formulas`, `total_errors`, and an
`error_summary` naming up to 100 cells per error type (`locations_truncated` says how many it
withheld — trust `total_errors`, not the length of the list). Fix what it names and run it
again. **JSON with an `error` key instead of a `status` means nothing was recalculated**, and
only that case exits non-zero — `errors_found` exits 0, so never treat a clean exit as a clean
workbook.

**A green recalc proves your formulas *evaluate*, not that they are *right*.** An off-by-one
range or a reference to the wrong row yields a clean, error-free file with wrong numbers.
Write 2–3 formulas first and check they pull the values you expect, before building out a grid.

**`not_computed` lists formulas LibreOffice cannot evaluate, even on `success`**: a function it lacks
(`_xlfn.XLOOKUP`, `_xlfn.LET`…) or a link to another workbook (`='[1]Returns Analysis'!$B$2`). They
keep what they held: after an openpyxl save, nothing until Excel opens it. For their numbers sooner,
read the original with `data_only=True`; keep the formulas. A sound formula depending on a listed one,
directly or not, reports `#NAME?` too: leave it and tell the user. A cell that tests one
(`IFERROR(B1,0)`) gets LibreOffice's fallback, and a kept value may be stale: say so. After editing
sheet XML, delete each changed formula's `<v>`: a mistake that keeps its value is listed, not reported.

## Choosing formulas that survive verification

LibreOffice implements fewer functions than Excel, and a formula it cannot evaluate has no value
until Excel opens the file.

- **Prefer Excel-2007-era functions** — `SUMIFS`, `INDEX`, `MATCH`, `IFERROR`, `SUMPRODUCT` — which need no prefix.
- **Six post-2007 functions work, but only with an `_xlfn.` prefix**, because openpyxl writes your formula into the XML verbatim and Excel stores post-2007 names prefixed (its UI hides the prefix): `_xlfn.TEXTJOIN`, `_xlfn.CONCAT`, `_xlfn.IFS`, `_xlfn.SWITCH`, `_xlfn.MAXIFS`, `_xlfn.MINIFS`. Written bare, each yields `#NAME?`.
- **Never use `XLOOKUP`, `XMATCH`, `SORT`, `FILTER`, `UNIQUE`, or `SEQUENCE`.** The runtime's LibreOffice cannot evaluate them, and one that can fills only the top-left cell of their result, since an openpyxl-written file has no spill metadata. Use `INDEX`/`MATCH` for lookups, and sort, filter, and de-duplicate in Python before writing the cells.
- A missing `)` is reported as `unbalanced brackets`. Too few arguments (`=ROUND(A1)`) is not: LibreOffice returns a value, and Excel may offer to repair the file.

## openpyxl gotchas

- **Reading a model takes two loads.** `data_only=True` yields cached values with the formulas gone; the default yields formula strings with no values. One pass cannot give you both.
- **`data_only=True` is destructive if you save.** That workbook has no formulas left, so saving replaces every one with a literal — permanently.
- **`data_only=True` on a file openpyxl just wrote returns `None` everywhere** — run `recalc.py` first. (A formula whose result is `""` also reads back as `None`.)
- **Merged cells: write the top-left anchor only.** Every other cell in the range is a `MergedCell` whose `.value` is read-only.
- **`.xlsm` loses its macros unless you pass `keep_vba=True`** to `load_workbook`.
- **A sheet name containing a space must be quoted** in a cross-sheet reference: `='Assumptions Inputs'!$B$5`. Unquoted, it evaluates to `#VALUE!`.
- **Saving an upload drops its text boxes, shapes and slicers.** First search `xl/drawings/` for `<xdr:sp ` (with the space) and look for an `xl/slicers/` part; if found, tell the user or edit the sheet XML directly.
- **Nothing mends a malformed workbook**: an unclosed quote in a number format (`'0.00 "USD'`), a defined name starting with `=`, or duplicate or empty table headers are saved as written, and Excel may offer to repair the file.
- **Partly bold or italic cell text is flattened** unless `load_workbook(rich_text=True)`; such a `.value` is then a `CellRichText`, not a `str`.
- **Every new formula needs its own `number_format`** (date, %, currency): nothing infers it, so a date shows as `45337` and 30% as `0.3`.
- **On every new chart set `chart.roundedCorners = False`, `.overlay = False` on its legend and on each title you gave it (`chart.legend`, `chart.title`, `chart.x_axis.title`…), and, if it has axes (not pie or doughnut), `chart.x_axis.delete = False` and `chart.y_axis.delete = False`**, or Excel may draw the title or legend over the plot, round the frame, or draw no axes.
- **With a `DateAxis`, set `chart.y_axis.crossAx = 500`**, or Excel may offer to repair the file.

## Financial models

Unless the user says otherwise, or the existing file already does something else.

**Color:** blue text (`0,0,255`) for hardcoded inputs and scenario levers · black for formulas ·
green (`0,128,0`) for links to another sheet · red (`255,0,0`) for links to another file ·
yellow fill (`255,255,0`) for key assumptions and cells the user should fill in.

**Numbers:** currency `$#,##0`, with the unit named in the header (`Revenue ($mm)`) · zeros
render as `-`, including in percentages (`$#,##0;($#,##0);-`) · negatives in parentheses ·
percentages `0.0%`, **stored as fractions** (`0.15` renders `15.0%`; storing `15` renders
`1500.0%`) · valuation multiples `0.0x` · years as text (`"2024"`, never `2,024`).

**Structure:** every assumption in its own labeled cell, referenced by the formulas that use it
(`=B5*(1+$B$6)`, never `=B5*1.05`) · formulas consistent across every projection period, since a
lone edited cell mid-row is the commonest silent error · guard denominators that can be zero.

## Dependencies

`openpyxl`, `pandas`, `markitdown` (pip, preinstalled — install only if an import fails or the command is missing) · LibreOffice (`soffice`, auto-configured for sandboxed environments via `scripts/office/soffice.py`)
