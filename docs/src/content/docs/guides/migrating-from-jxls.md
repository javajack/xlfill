---
title: "Migrating from JXLS to XLFill"
description: "Step-by-step guide to moving your JXLS 3.0 templates to XLFill — same jx: syntax, more features, Go performance."
---

If you're coming from JXLS, you're going to feel right at home. XLFill is a Go port of the JXLS 3.0 template engine — same `jx:` command syntax, same cell comment approach, same template philosophy. Your existing templates will work with minimal changes.

The difference? XLFill adds 12 commands JXLS doesn't have, fixes long-standing JXLS pain points, and runs as a native Go binary — no JVM, no Maven, no Spring Boot.

## Syntax comparison: nearly identical

| Feature | JXLS 3.0 | XLFill |
|---------|----------|--------|
| Area definition | `jx:area(lastCell="C2")` | `jx:area(lastCell="C2")` |
| Loop | `jx:each(items="employees" var="e" lastCell="C2")` | `jx:each(items="employees" var="e" lastCell="C2")` |
| Conditional | `jx:if(condition="e.salary > 50000" lastCell="C2")` | `jx:if(condition="e.Salary > 50000" lastCell="C2")` |
| Expressions | `${e.name}` | `${e.Name}` |
| Grid layout | `jx:grid(headers="headers" data="data" lastCell="A2")` | `jx:grid(headers="headers" data="data" lastCell="A2")` |
| Image | `jx:image(src="logoBytes" lastCell="C5")` | `jx:image(src="logoBytes" lastCell="C5")` |
| Cell merging | `jx:mergeCells(lastCell="C2")` | `jx:mergeCells(lastCell="C2")` |

The commands are literally the same. The only consistent difference is Go's convention of capitalized exported field names (`e.Name` vs `e.name`).

> **Pro tip:** If your JXLS templates use lowercase field names like `${e.name}`, you can keep them — just use `map[string]any` in Go with lowercase keys. No template changes needed.

## What XLFill adds (12 commands JXLS lacks)

| Command | What it does | JXLS equivalent |
|---------|-------------|-----------------|
| `jx:table` | Auto-filter tables with banding | None — requires POI post-processing |
| `jx:chart` | Bar, line, pie, and 5 more chart types | None — requires POI post-processing |
| `jx:sparkline` | Inline trend charts in cells | None |
| `jx:conditionalFormat` | Data bars, color scales, icon sets | None — JXLS issue #152 |
| `jx:dataValidation` | Dropdowns, number/date constraints | None — JXLS issue #46 |
| `jx:definedName` | Named ranges for downstream formulas | None |
| `jx:group` | Collapsible row/column groups | None |
| `jx:autoRowHeight` | Auto-fit row heights | None |
| `jx:autoColWidth` | Auto-fit column widths | None |
| `jx:freezePanes` | Freeze header rows/columns | None |
| `jx:pageBreak` | Page breaks for printing | None |
| `jx:protect` | Sheet protection with password | None |

Every one of these is a zero-code template command. In JXLS, you'd need to manipulate the output file with Apache POI after JXLS runs — writing Java code, managing POI dependencies, and hoping the post-processing doesn't break JXLS's output.

## Migration steps

### Step 1: Copy your .xlsx template

Your JXLS template works as-is with XLFill. Copy it into your Go project.

```bash
cp /java-project/src/main/resources/templates/report.xlsx /go-project/templates/
```

If your templates use lowercase field names (`${e.name}`), they'll work fine with `map[string]any` data. If you want to use Go structs, update the expressions to match struct field names (`${e.Name}`).

### Step 2: Replace Java data with Go map

**JXLS (Java):**
```java
Context context = new Context();
context.putVar("employees", employeeList);
context.putVar("title", "Monthly Report");
JxlsHelper.getInstance().processTemplate(
    new FileInputStream("template.xlsx"),
    new FileOutputStream("output.xlsx"),
    context
);
```

**XLFill (Go):**
```go
data := map[string]any{
    "employees": employees,
    "title":     "Monthly Report",
}
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

Same concept, fewer lines, no context object ceremony.

### Step 3: Convert Java collections to Go data

XLFill accepts `map[string]any` with slices of maps or structs. Here's how common JXLS data patterns translate:

**Java List<Map>** becomes **Go slice of maps:**
```go
data := map[string]any{
    "employees": []map[string]any{
        {"name": "Alice", "department": "Engineering", "salary": 95000},
        {"name": "Bob", "department": "Sales", "salary": 72000},
    },
}
```

**Java List<Bean>** becomes **Go slice of structs:**
```go
type Employee struct {
    Name       string
    Department string
    Salary     float64
}

data := xlfill.StructSliceToData("employees", []Employee{
    {Name: "Alice", Department: "Engineering", Salary: 95000},
    {Name: "Bob", Department: "Sales", Salary: 72000},
})
```

> **Pro tip:** `StructSliceToData` converts struct fields to map entries automatically. It handles nested structs, so `employee.Address.City` works in templates just like JXLS.

**JSON from an API** — if your Java app was calling a REST API and feeding the JSON into JXLS, XLFill has a one-liner:

```go
data, err := xlfill.JSONToData(jsonBytes)
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

**Database queries** — if your JXLS data came from JDBC:

```go
rows, _ := db.Query("SELECT name, department, salary FROM employees")
data, err := xlfill.SQLRowsToData("employees", rows)
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

### Step 4: Run it

```bash
go run main.go
```

Open the output. Compare it with JXLS output. It should be identical for the shared command set.

## Feature gap table: JXLS vs XLFill

| Feature | JXLS 3.0 | XLFill |
|---------|----------|--------|
| Template commands | 8 | **20** |
| Built-in expression functions | 0 | **18** |
| Template validation (pre-fill) | No | **Yes** |
| Data contract validation | No | **Yes** |
| Streaming mode | No | **3x speed, 60% less memory** |
| Parallel processing | No | **Built-in** |
| Auto-mode optimization | No | **Built-in** |
| Compiled/cached templates | No | **Built-in** |
| Batch generation | No | **FillBatch** |
| HTTP handler | No | **Built-in** |
| Progress reporting | No | **Built-in** |
| Context cancellation | No | **Built-in** |
| Structured error messages | No | **Yes (kind + cell + message)** |
| "Did you mean?" suggestions | No | **Yes** |
| Custom expression delimiters | No | **Yes** |
| Debug trace output | No | **Yes** |
| Loop filter/sort/group | Partial (custom Java) | **Built-in attributes** |
| Multisheet from single loop | Partial | **Built-in** |

## Common JXLS pain points solved

### Data validation in loops (JXLS issue #46)

JXLS can't preserve data validation (dropdowns) when rows expand. If you had a dropdown in your template row, it disappeared in the output. XLFill's `jx:dataValidation` command explicitly handles this:

```
Cell A2 comment:
  jx:each(items="employees" var="e" lastCell="C2")

Cell C2 comment:
  jx:dataValidation(type="list" formula="Active,Inactive,On Leave" lastCell="C2")

Cell C2: ${e.Status}
```

Every output row gets the dropdown. No post-processing needed.

### Conditional formatting performance (JXLS issue #152)

JXLS had no built-in support for conditional formatting. Users resorted to Apache POI post-processing, which was slow and fragile. XLFill handles it natively:

```
Cell B2 comment:
  jx:conditionalFormat(lastCell="B2" type="colorScale" minColor="FF0000" maxColor="00FF00")
```

Every salary cell gets a color scale — red for low, green for high. Zero Java code, zero POI manipulation.

### Row grouping (JXLS issue #250)

Collapsible row groups were a frequent JXLS feature request with no official support. XLFill's `jx:group` command handles it:

```
Cell A2 comment:
  jx:each(items="departments" var="dept" lastCell="C10")
  jx:group(lastCell="C10" collapsed="false")
```

Department sections become collapsible outline groups in the output.

## Expression language differences

JXLS uses JEXL (Java Expression Language). XLFill uses [expr-lang](https://github.com/expr-lang/expr), a fast Go expression evaluator. Most expressions work identically, but there are a few differences:

| Expression | JXLS (JEXL) | XLFill (expr-lang) |
|-----------|-------------|---------------------|
| Null-safe access | `e.?name` | `e.Name` (nil-safe by default) |
| String concatenation | `e.first + ' ' + e.last` | `e.First + " " + e.Last` |
| Ternary | `e.active ? 'Yes' : 'No'` | `e.Active ? "Yes" : "No"` |
| Method calls | `e.getFullName()` | Not supported — use functions |
| Built-in functions | None | 18 built-in: `upper()`, `formatNumber()`, `sumBy()`, etc. |

> **Gotcha:** JXLS allows calling Java methods in expressions (`e.getFullName()`). XLFill doesn't — Go structs don't have the same reflection model. Instead, either compute the value before filling, or register a custom function with `WithFunction`.

## JXLS custom transformers → XLFill equivalents

| JXLS pattern | XLFill equivalent |
|-------------|-------------------|
| Custom Transformer | `WithCommand(name, factory)` for new commands |
| AreaListener (Java) | `WithAreaListener(listener)` — same concept, Go interface |
| Custom function in JEXL | `WithFunction(name, fn)` |
| Post-processing with POI | `WithPreWrite(fn)` callback, or post-process with excelize |

## Tips for a smooth migration

- **Start with one template.** Pick your simplest JXLS template, migrate it, and verify the output. Then move to complex ones.
- **Use `Validate()` liberally.** XLFill can check your template for syntax errors without data — JXLS couldn't do this. Run it after copying every template.
- **Leverage new features gradually.** Once the basic migration works, add `jx:chart`, `jx:table`, `jx:conditionalFormat` to templates that needed POI post-processing. Delete that Java code.
- **Test with `WithStrictMode(true)`.** This turns unknown command warnings into errors — catches typos and leftover JXLS-only syntax.
- **Batch convert templates.** If you have dozens of templates, write a small Go test that calls `Validate()` on each one — instant migration verification.

```go
func TestAllTemplates(t *testing.T) {
    templates, _ := filepath.Glob("templates/*.xlsx")
    for _, tmpl := range templates {
        issues, err := xlfill.Validate(tmpl)
        if err != nil {
            t.Errorf("%s: %v", tmpl, err)
        }
        for _, issue := range issues {
            t.Errorf("%s: %s", tmpl, issue)
        }
    }
}
```

## What's next?

- See the full feature comparison: **[Feature Comparison &rarr;](/xlfill/reference/comparison/)**
- Learn about XLFill's streaming and parallel modes (things JXLS can't do): **[Performance Tuning &rarr;](/xlfill/guides/performance-tuning/)**
- Build your first XLFill template: **[Getting Started &rarr;](/xlfill/guides/getting-started/)**
