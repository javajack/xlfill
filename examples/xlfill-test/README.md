# XLFill Feature Test Suite

A self-contained Go project that exercises every XLFill feature in one run. It programmatically creates template `.xlsx` files (saved to `input/`), fills them using `xlfill.Fill` / `FillBytes` / `FillReader`, writes the results to `output/`, and verifies correctness.

## Quick Start

```bash
cd examples/xlfill-test
go run .
```

Expected output:

```
01_basic_each                       OK
02_each_with_varindex               OK
03_each_direction_right             OK
...
19_fill_reader                      OK

19 passed, 0 failed out of 19 tests
```

## What's Tested

| # | Test | Feature | Template | Output |
|---|------|---------|----------|--------|
| 01 | `basic_each` | `jx:each` — iterate a list downward | `input/t01.xlsx` | `output/01_basic_each.xlsx` |
| 02 | `each_with_varindex` | `varIndex` — access loop index | `input/t02.xlsx` | `output/02_varindex.xlsx` |
| 03 | `each_direction_right` | `direction="RIGHT"` — horizontal expansion | `input/t03.xlsx` | `output/03_direction_right.xlsx` |
| 04 | `each_with_select` | `select` — filter items by expression | `input/t04.xlsx` | `output/04_select.xlsx` |
| 05 | `each_with_orderby` | `orderBy` — sort items | `input/t05.xlsx` | `output/05_orderby.xlsx` |
| 06 | `each_with_groupby` | `groupBy` — group items into `GroupData` | `input/t06.xlsx` | `output/06_groupby.xlsx` |
| 07 | `if_command` | `jx:if` — conditional cell rendering | `input/t07.xlsx` | `output/07_if_command.xlsx` |
| 08 | `formulas` | Formula auto-expansion after `each` grows rows | `input/t08.xlsx` | `output/08_formulas.xlsx` |
| 09 | `grid_command` | `jx:grid` — dynamic header + data grid | `input/t09.xlsx` | `output/09_grid.xlsx` |
| 10 | `image_command` | `jx:image` — embed PNG image from `[]byte` | `input/t10.xlsx` | `output/10_image.xlsx` |
| 11 | `merge_cells` | `jx:mergeCells` — dynamic cell merging | `input/t11.xlsx` | `output/11_mergecells.xlsx` |
| 12 | `hyperlinks` | `hyperlink()` — clickable links in cells | `input/t12.xlsx` | `output/12_hyperlinks.xlsx` |
| 13 | `nested_each` | Nested `jx:each` — departments with employees | `input/t13.xlsx` | `output/13_nested_each.xlsx` |
| 14 | `multisheet` | `multisheet` — one sheet per item | `input/t14.xlsx` | `output/14_multisheet.xlsx` |
| 15 | `custom_notation` | `WithExpressionNotation("{{", "}}")` | `input/t15.xlsx` | `output/15_custom_notation.xlsx` |
| 16 | `keep_template_sheet` | `WithKeepTemplateSheet(true)` | `input/t16.xlsx` | `output/16_keep_template.xlsx` |
| 17 | `autorowheight` | `jx:autoRowHeight` — auto-fit row height | `input/t17.xlsx` | `output/17_autorowheight.xlsx` |
| 18 | `fill_bytes` | `xlfill.FillBytes` API | `input/t18.xlsx` | `output/18_fill_bytes.xlsx` |
| 19 | `fill_reader` | `xlfill.FillReader` API (io.Reader/Writer) | `input/t19.xlsx` | `output/19_fill_reader.xlsx` |

## Project Structure

```
examples/xlfill-test/
  main.go        Single-file test runner
  go.mod         Module with local replace directive
  go.sum
  README.md
  input/         Generated template files (open these to see jx: comments)
  output/        Generated output files (open these to see filled results)
```

## How It Works

Each test function:

1. Creates a template `.xlsx` using [excelize](https://github.com/xuri/excelize) — sets cell values with `${...}` expressions and adds `jx:` commands as cell comments
2. Saves the template to `input/`
3. Calls `xlfill.Fill()` (or `FillBytes` / `FillReader`) to process it
4. Saves the result to `output/`
5. Opens the output and asserts specific cell values

You can open any `input/*.xlsx` file in Excel, Google Sheets, or LibreOffice to inspect the template structure (look at cell comments). Then open the matching `output/*.xlsx` to see the filled result.
