---
title: "Examples"
description: Runnable examples for every XLFill feature, with template inputs and filled outputs.
---

The [`examples/xlfill-test`](https://github.com/javajack/xlfill/tree/main/examples/xlfill-test) project in the repository exercises every XLFill feature in one Go program. Each test creates a template, fills it, and verifies the output.

You can run it yourself:

```bash
cd examples/xlfill-test
go run .
```

Or just browse the files below. Open any **input** file in Excel or LibreOffice to see the `jx:` comments and `${...}` expressions. Then open the matching **output** file to see the filled result.

## All examples

| # | Feature | Command / API | Template | Output | Source |
|---|---------|---------------|----------|--------|--------|
| 01 | Basic loop | [`jx:each`](/xlfill/commands/each/) | [t01.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t01.xlsx) | [01_basic_each.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/01_basic_each.xlsx) | [code](#01-basic-loop) |
| 02 | Loop index | [`jx:each` varIndex](/xlfill/commands/each/#iteration-index) | [t02.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t02.xlsx) | [02_varindex.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/02_varindex.xlsx) | [code](#02-loop-index) |
| 03 | Expand RIGHT | [`jx:each` direction](/xlfill/commands/each/#expand-right) | [t03.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t03.xlsx) | [03_direction_right.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/03_direction_right.xlsx) | [code](#03-expand-right) |
| 04 | Filter items | [`jx:each` select](/xlfill/commands/each/#filtering) | [t04.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t04.xlsx) | [04_select.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/04_select.xlsx) | [code](#04-filter-items) |
| 05 | Sort items | [`jx:each` orderBy](/xlfill/commands/each/#sorting) | [t05.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t05.xlsx) | [05_orderby.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/05_orderby.xlsx) | [code](#05-sort-items) |
| 06 | Group items | [`jx:each` groupBy](/xlfill/commands/each/#grouping) | [t06.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t06.xlsx) | [06_groupby.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/06_groupby.xlsx) | [code](#06-group-items) |
| 07 | Conditional | [`jx:if`](/xlfill/commands/if/) | [t07.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t07.xlsx) | [07_if_command.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/07_if_command.xlsx) | [code](#07-conditional) |
| 08 | Formulas | [Formula expansion](/xlfill/guides/formulas/) | [t08.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t08.xlsx) | [08_formulas.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/08_formulas.xlsx) | [code](#08-formulas) |
| 09 | Dynamic grid | [`jx:grid`](/xlfill/commands/grid/) | [t09.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t09.xlsx) | [09_grid.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/09_grid.xlsx) | [code](#09-dynamic-grid) |
| 10 | Embed image | [`jx:image`](/xlfill/commands/image/) | [t10.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t10.xlsx) | [10_image.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/10_image.xlsx) | [code](#10-embed-image) |
| 11 | Merge cells | [`jx:mergeCells`](/xlfill/commands/mergecells/) | [t11.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t11.xlsx) | [11_mergecells.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/11_mergecells.xlsx) | [code](#11-merge-cells) |
| 12 | Hyperlinks | [`hyperlink()` expression](/xlfill/guides/expressions/) | [t12.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t12.xlsx) | [12_hyperlinks.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/12_hyperlinks.xlsx) | [code](#12-hyperlinks) |
| 13 | Nested loops | [Nested `jx:each`](/xlfill/commands/each/#nested-loops) | [t13.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t13.xlsx) | [13_nested_each.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/13_nested_each.xlsx) | [code](#13-nested-loops) |
| 14 | Multi-sheet | [`jx:each` multisheet](/xlfill/commands/each/#multi-sheet-output) | [t14.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t14.xlsx) | [14_multisheet.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/14_multisheet.xlsx) | [code](#14-multi-sheet) |
| 15 | Custom notation | [`WithExpressionNotation`](/xlfill/reference/api/) | [t15.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t15.xlsx) | [15_custom_notation.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/15_custom_notation.xlsx) | [code](#15-custom-notation) |
| 16 | Keep template sheet | [`WithKeepTemplateSheet`](/xlfill/reference/api/) | [t16.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t16.xlsx) | [16_keep_template.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/16_keep_template.xlsx) | [code](#16-keep-template-sheet) |
| 17 | Auto row height | [`jx:autoRowHeight`](/xlfill/commands/autorowheight/) | [t17.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t17.xlsx) | [17_autorowheight.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/17_autorowheight.xlsx) | [code](#17-auto-row-height) |
| 18 | FillBytes API | [`xlfill.FillBytes`](/xlfill/reference/api/) | [t18.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t18.xlsx) | [18_fill_bytes.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/18_fill_bytes.xlsx) | [code](#18-fillbytes-api) |
| 19 | FillReader API | [`xlfill.FillReader`](/xlfill/reference/api/) | [t19.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t19.xlsx) | [19_fill_reader.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/19_fill_reader.xlsx) | [code](#19-fillreader-api) |

---

## Code snippets

### 01 — Basic loop

Template cells: `A1:"Name"` `B1:"Age"` `C1:"Salary"` `A2:"${e.Name}"` `B2:"${e.Age}"` `C2:"${e.Salary}"`

```
Cell A1 comment:  jx:area(lastCell="C2")
Cell A2 comment:  jx:each(items="employees" var="e" lastCell="C2")
```

```go
data := map[string]any{
    "employees": []any{
        map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000},
        map[string]any{"Name": "Bob", "Age": 25, "Salary": 6000},
        map[string]any{"Name": "Carol", "Age": 35, "Salary": 7000},
    },
}
xlfill.Fill("input/t01.xlsx", "output/01_basic_each.xlsx", data)
```

### 02 — Loop index

Template cells: `A1:"#"` `B1:"Item"` `A2:"${idx + 1}"` `B2:"${e}"`

```
Cell A1 comment:  jx:area(lastCell="B2")
Cell A2 comment:  jx:each(items="items" var="e" varIndex="idx" lastCell="B2")
```

```go
data := map[string]any{"items": []any{"Apple", "Banana", "Cherry"}}
xlfill.Fill("input/t02.xlsx", "output/02_varindex.xlsx", data)
```

### 03 — Expand RIGHT

Template cells: `A1:"${e}"`

```
Cell A1 comment:
  jx:area(lastCell="A1")
  jx:each(items="months" var="e" direction="RIGHT" lastCell="A1")
```

```go
data := map[string]any{"months": []any{"Jan", "Feb", "Mar", "Apr"}}
xlfill.Fill("input/t03.xlsx", "output/03_direction_right.xlsx", data)
```

### 04 — Filter items

Template cells: `A1:"Name"` `B1:"Salary"` `A2:"${e.Name}"` `B2:"${e.Salary}"`

```
Cell A1 comment:  jx:area(lastCell="B2")
Cell A2 comment:  jx:each(items="employees" var="e" select="e.Salary >= 6000" lastCell="B2")
```

```go
data := map[string]any{
    "employees": []any{
        map[string]any{"Name": "Alice", "Salary": 5000},
        map[string]any{"Name": "Bob", "Salary": 6000},
        map[string]any{"Name": "Carol", "Salary": 7000},
    },
}
// Output contains only Bob and Carol
xlfill.Fill("input/t04.xlsx", "output/04_select.xlsx", data)
```

### 05 — Sort items

Template cells: `A1:"Name"` `A2:"${e.Name}"`

```
Cell A1 comment:  jx:area(lastCell="A2")
Cell A2 comment:  jx:each(items="names" var="e" orderBy="e.Name DESC" lastCell="A2")
```

```go
data := map[string]any{
    "names": []any{
        map[string]any{"Name": "Charlie"},
        map[string]any{"Name": "Alice"},
        map[string]any{"Name": "Bob"},
    },
}
// Output order: Charlie, Bob, Alice
xlfill.Fill("input/t05.xlsx", "output/05_orderby.xlsx", data)
```

### 06 — Group items

Template cells: `A1:"${g.Item.Department}"` `B1:"${g.Item.Name}"`

```
Cell A1 comment:
  jx:area(lastCell="B1")
  jx:each(items="employees" var="g" groupBy="g.Department" lastCell="B1")
```

```go
data := map[string]any{
    "employees": []any{
        map[string]any{"Name": "Alice", "Department": "Engineering"},
        map[string]any{"Name": "Bob", "Department": "Sales"},
        map[string]any{"Name": "Carol", "Department": "Engineering"},
    },
}
// Output: 2 rows — Engineering, Sales
xlfill.Fill("input/t06.xlsx", "output/06_groupby.xlsx", data)
```

### 07 — Conditional

Template cells: `A1:"Name"` `B1:"Status"` `A2:"${e.Name}"` `B2:"ACTIVE"`

```
Cell A1 comment:  jx:area(lastCell="B2")
Cell A2 comment:  jx:each(items="employees" var="e" lastCell="B2")
Cell B2 comment:  jx:if(condition="e.Active" lastCell="B2")
```

```go
data := map[string]any{
    "employees": []any{
        map[string]any{"Name": "Alice", "Active": true},
        map[string]any{"Name": "Bob", "Active": false},
        map[string]any{"Name": "Carol", "Active": true},
    },
}
xlfill.Fill("input/t07.xlsx", "output/07_if_command.xlsx", data)
```

### 08 — Formulas

Template cells: `A1:"Amount"` `A2:"${e.Amount}"` `A3: =SUM(A2:A2)`

```
Cell A1 comment:  jx:area(lastCell="A3")
Cell A2 comment:  jx:each(items="items" var="e" lastCell="A2")
```

```go
data := map[string]any{
    "items": []any{
        map[string]any{"Amount": 100},
        map[string]any{"Amount": 200},
        map[string]any{"Amount": 300},
    },
}
// SUM formula auto-expands to =SUM(A2:A4)
xlfill.Fill("input/t08.xlsx", "output/08_formulas.xlsx", data)
```

### 09 — Dynamic grid

Template cells: `A1:"placeholder"`

```
Cell A1 comment:
  jx:area(lastCell="A2")
  jx:grid(headers="headers" data="data" lastCell="A2")
```

```go
data := map[string]any{
    "headers": []any{"Name", "Age", "City"},
    "data": []any{
        []any{"Alice", 30, "NYC"},
        []any{"Bob", 25, "LA"},
    },
}
xlfill.Fill("input/t09.xlsx", "output/09_grid.xlsx", data)
```

### 10 — Embed image

Template cells: `A1:"Logo below"` `A2:""`

```
Cell A1 comment:  jx:area(lastCell="A2")
Cell A2 comment:  jx:image(src="logo" imageType="PNG" lastCell="A2")
```

```go
logoBytes, _ := os.ReadFile("logo.png")
data := map[string]any{"logo": logoBytes}
xlfill.Fill("input/t10.xlsx", "output/10_image.xlsx", data)
```

### 11 — Merge cells

Template cells: `A1:"Merged Header"`

```
Cell A1 comment:
  jx:area(lastCell="C2")
  jx:mergeCells(lastCell="C2" cols="3" rows="2")
```

```go
data := map[string]any{}
xlfill.Fill("input/t11.xlsx", "output/11_mergecells.xlsx", data)
```

### 12 — Hyperlinks

Template cells: `A1:"Site"` `B1:"Link"` `A2:"${e.Name}"` `B2:"${hyperlink(e.URL, e.Name)}"`

```
Cell A1 comment:  jx:area(lastCell="B2")
Cell A2 comment:  jx:each(items="sites" var="e" lastCell="B2")
```

```go
data := map[string]any{
    "sites": []any{
        map[string]any{"Name": "Google", "URL": "https://google.com"},
        map[string]any{"Name": "GitHub", "URL": "https://github.com"},
    },
}
xlfill.Fill("input/t12.xlsx", "output/12_hyperlinks.xlsx", data)
```

### 13 — Nested loops

Template cells: `A1:"${dept.Name}"` `A2:"${e.Name}"` `B2:"${e.Role}"`

```
Cell A1 comment:
  jx:area(lastCell="B2")
  jx:each(items="departments" var="dept" lastCell="B2")
Cell A2 comment:
  jx:each(items="dept.Employees" var="e" lastCell="B2")
```

```go
data := map[string]any{
    "departments": []any{
        map[string]any{
            "Name": "Engineering",
            "Employees": []any{
                map[string]any{"Name": "Alice", "Role": "Lead"},
                map[string]any{"Name": "Bob", "Role": "Dev"},
            },
        },
        map[string]any{
            "Name": "Sales",
            "Employees": []any{
                map[string]any{"Name": "Carol", "Role": "Manager"},
            },
        },
    },
}
xlfill.Fill("input/t13.xlsx", "output/13_nested_each.xlsx", data)
```

### 14 — Multi-sheet

Template cells: `A1:"${dept.Name}"` `A2:"${dept.Head}"`

```
Cell A1 comment:
  jx:area(lastCell="A2")
  jx:each(items="departments" var="dept" multisheet="sheetNames" lastCell="A2")
```

```go
data := map[string]any{
    "sheetNames":  []any{"Engineering", "Sales", "HR"},
    "departments": []any{
        map[string]any{"Name": "Engineering", "Head": "Alice"},
        map[string]any{"Name": "Sales", "Head": "Bob"},
        map[string]any{"Name": "HR", "Head": "Carol"},
    },
}
// Creates 3 sheets: Engineering, Sales, HR
xlfill.Fill("input/t14.xlsx", "output/14_multisheet.xlsx", data)
```

### 15 — Custom notation

Template cells: `A1:"Name"` `A2:"{{e.Name}}"` (uses `{{ }}` instead of `${ }`)

```
Cell A1 comment:  jx:area(lastCell="A2")
Cell A2 comment:  jx:each(items="items" var="e" lastCell="A2")
```

```go
data := map[string]any{
    "items": []any{
        map[string]any{"Name": "Alpha"},
        map[string]any{"Name": "Beta"},
    },
}
xlfill.Fill("input/t15.xlsx", "output/15_custom_notation.xlsx", data,
    xlfill.WithExpressionNotation("{{", "}}"),
)
```

### 16 — Keep template sheet

```go
data := map[string]any{"items": []any{"X", "Y"}}
xlfill.Fill("input/t16.xlsx", "output/16_keep_template.xlsx", data,
    xlfill.WithKeepTemplateSheet(true),
)
```

### 17 — Auto row height

Template cells: `A1:"${text}"`

```
Cell A1 comment:
  jx:area(lastCell="A1")
  jx:autoRowHeight(lastCell="A1")
```

```go
data := map[string]any{
    "text": "This is a long text that should cause the row height to be adjusted automatically.",
}
xlfill.Fill("input/t17.xlsx", "output/17_autorowheight.xlsx", data)
```

### 18 — FillBytes API

```go
data := map[string]any{"items": []any{"One", "Two", "Three"}}
outBytes, err := xlfill.FillBytes("input/t18.xlsx", data)
os.WriteFile("output/18_fill_bytes.xlsx", outBytes, 0o644)
```

### 19 — FillReader API

```go
tmplBytes, _ := os.ReadFile("input/t19.xlsx")
data := map[string]any{"items": []any{"Red", "Green", "Blue"}}
var out bytes.Buffer
xlfill.FillReader(bytes.NewReader(tmplBytes), &out, data)
os.WriteFile("output/19_fill_reader.xlsx", out.Bytes(), 0o644)
```
