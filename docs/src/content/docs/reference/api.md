---
title: API Reference
description: Complete API reference for the XLFill Go library — functions, options, and types.
---

This page covers every public function, option, and type in the XLFill library. For guided walkthroughs, see the [Getting Started](/xlfill/guides/getting-started/) guide.

## Top-level functions

These are the simplest way to use XLFill. One function call, done.

### Fill

```go
func Fill(templatePath, outputPath string, data map[string]any, opts ...Option) error
```

Read a template file, fill it with data, write the output file.

```go
xlfill.Fill("template.xlsx", "report.xlsx", data)
```

### FillBytes

```go
func FillBytes(templatePath string, data map[string]any, opts ...Option) ([]byte, error)
```

Read a template file, fill it, return the result as bytes. Useful when you need the output in memory.

```go
bytes, err := xlfill.FillBytes("template.xlsx", data)
```

### FillReader

```go
func FillReader(template io.Reader, output io.Writer, data map[string]any, opts ...Option) error
```

Fill from an `io.Reader`, write to an `io.Writer`. Perfect for HTTP handlers — no temp files needed:

```go
func handler(w http.ResponseWriter, r *http.Request) {
    tmpl, _ := os.Open("template.xlsx")
    defer tmpl.Close()
    xlfill.FillReader(tmpl, w, data)
}
```

### Validate

```go
func Validate(templatePath string, opts ...Option) ([]ValidationIssue, error)
```

Check a template for structural and expression errors **without requiring data**. Returns a list of issues found. Use this in CI pipelines or during development to catch problems early.

```go
issues, err := xlfill.Validate("template.xlsx")
if err != nil {
    log.Fatal(err) // template couldn't be opened or parsed at all
}
for _, issue := range issues {
    fmt.Println(issue) // [ERROR] Sheet1!B2: invalid expression syntax "e.Name +": ...
}
```

What it checks:
- **Expression syntax** — validates all `${...}` in cell values and formulas
- **Command attributes** — validates `items`, `condition`, `select`, `headers`, `data`, `count` expressions
- **Bounds** — verifies each command's `lastCell` fits within its parent area

### ValidateData

```go
func ValidateData(templatePath string, data map[string]any, opts ...Option) ([]ValidationIssue, error)
```

Verify that a data map satisfies the template's expression requirements. Extracts all `${...}` variable references, cross-references against the data map and command-provided variables, and reports mismatches.

```go
issues, err := xlfill.ValidateData("template.xlsx", data)
// [WARN] Sheet1!A1: expression ${companyName} references variable "companyName" which is not in data
```

See the [Error Handling guide](/xlfill/guides/error-handling/) for details.

### Describe

```go
func Describe(templatePath string, opts ...Option) (string, error)
```

Parse a template and return a human-readable tree showing the area hierarchy, commands with attributes, and expressions found in cells.

```go
output, err := xlfill.Describe("template.xlsx")
fmt.Print(output)
```

Sample output:
```
Template: template.xlsx
Sheet1!A1:C2 area (3x2)
  Commands:
    Sheet1!A2 each (3x1) items="employees" var="e"
      Sheet1!A2:C2 area (3x1)
        Expressions:
          A2: ${e.Name}
          B2: ${e.Age}
          C2: ${e.Salary}
```

### Compile

```go
func Compile(templatePath string, opts ...Option) (*CompiledTemplate, error)
```

Parse a template file once and return a reusable `CompiledTemplate`. Each subsequent fill reuses the cached template bytes — no file I/O. Ideal for batch generation and server-side endpoints.

```go
compiled, err := xlfill.Compile("template.xlsx", xlfill.WithRecalculateOnOpen(true))

// Reuse for multiple data sets
compiled.Fill(data1, "report_jan.xlsx")
compiled.Fill(data2, "report_feb.xlsx")
bytes, _ := compiled.FillBytes(data3)
```

### SuggestMode

```go
func SuggestMode(templatePath string, dataHint map[string]any, opts ...Option) (*ModeSuggestion, error)
```

Analyze a template and return the optimal processing mode (sequential, streaming, or parallel) based on its structure and a data size hint. See [Performance Tuning](/xlfill/guides/performance-tuning/).

```go
s, _ := xlfill.SuggestMode("template.xlsx", map[string]any{"itemCount": 50000})
fmt.Println(s.Mode, s.Reasons) // "streaming" ["large dataset (>=10K items)", ...]
```

### HTTPHandler

```go
func HTTPHandler(templatePath string, dataFunc func(r *http.Request) (map[string]any, error), opts ...Option) http.Handler
```

Serve Excel reports directly from an HTTP endpoint. The `dataFunc` extracts data from the request (query params, database lookup, etc.) and XLFill generates the .xlsx response with correct content headers.

```go
http.Handle("/report", xlfill.HTTPHandler("template.xlsx",
    func(r *http.Request) (map[string]any, error) {
        dept := r.URL.Query().Get("dept")
        employees, err := db.GetEmployeesByDept(dept)
        return map[string]any{"employees": employees}, err
    },
    xlfill.WithStreaming(true),
))
```

One handler, no temp files, no manual `Content-Type` headers. The response streams directly to the client.

## Filler (advanced)

For repeated fills or fine-grained control, create a `Filler`:

```go
filler := xlfill.NewFiller(
    xlfill.WithTemplate("template.xlsx"),
    xlfill.WithStrictMode(true),
    xlfill.WithRecalculateOnOpen(true),
)

err := filler.Fill(data, "output.xlsx")

// Check warnings (unknown commands, etc.)
for _, w := range filler.Warnings() {
    fmt.Println(w)
}

// Validate, ValidateData, Describe also available on Filler
issues, err := filler.Validate()
dataIssues, err := filler.ValidateData(data)
description, err := filler.Describe()
```

## Options

All options work with both the top-level functions and `NewFiller`.

### Template source

| Option | Description |
|--------|-------------|
| `WithTemplate(path)` | Set template file path |
| `WithTemplateReader(r io.Reader)` | Set template from a reader |

### Expression configuration

| Option | Description |
|--------|-------------|
| `WithExpressionNotation(begin, end)` | Custom delimiters (default: `${`, `}`) |

### Output control

| Option | Description |
|--------|-------------|
| `WithClearTemplateCells(bool)` | Clear unexpanded `${...}` cells (default: `true`) |
| `WithKeepTemplateSheet(bool)` | Keep template sheet in output (default: `false`) |
| `WithHideTemplateSheet(bool)` | Hide template sheet instead of removing (default: `false`) |
| `WithRecalculateOnOpen(bool)` | Tell Excel to recalculate formulas on open |

### Performance and scaling

| Option | Description |
|--------|-------------|
| `WithStreaming(bool)` | Use StreamWriter for output — 3x faster, 60% less memory. [Details &rarr;](/xlfill/guides/performance-tuning/) |
| `WithParallelism(n int)` | Concurrent `jx:each` processing with N goroutines. [Details &rarr;](/xlfill/guides/performance-tuning/) |
| `WithAutoMode(dataHint)` | Auto-select optimal mode from template analysis. [Details &rarr;](/xlfill/guides/performance-tuning/) |
| `WithContext(ctx)` | Set `context.Context` for cancellation and timeouts |
| `WithProgressFunc(fn)` | Progress callback: `func(FillProgress)` called per row |

### Quality and debugging

| Option | Description |
|--------|-------------|
| `WithStrictMode(bool)` | Turn unknown command warnings into errors. [Details &rarr;](/xlfill/guides/error-handling/) |
| `WithDebugWriter(w io.Writer)` | Structured trace output during processing. [Details &rarr;](/xlfill/guides/debugging/) |

### Selective streaming

| Option | Description |
|--------|-------------|
| `WithStreamingSheets(sheets...)` | Enable streaming only for specific sheets (by name). Other sheets use sequential mode. |

### Document metadata

| Option | Description |
|--------|-------------|
| `WithDocumentProperties(props)` | Set workbook properties (title, author, subject, description, etc.) |

### Extensibility

| Option | Description |
|--------|-------------|
| `WithCommand(name, factory)` | Register a [custom command](/xlfill/guides/custom-commands/) |
| `WithFunction(name, fn)` | Register a [custom expression function](/xlfill/guides/built-in-functions/) |
| `WithI18n(translations)` | Provide translation map for the `t()` function |
| `WithAreaListener(listener)` | Add a [cell transform hook](/xlfill/guides/area-listeners/) |
| `WithPreWrite(fn)` | Callback before writing output |

## Types

### Error types

```go
type ErrorKind int
const (
    ErrTemplate ErrorKind = iota  // template structure problem
    ErrData                       // data/expression problem
    ErrRuntime                    // I/O or system problem
)

type XLFillError struct {
    Kind    ErrorKind
    Cell    CellRef
    Command string
    Message string
    Err     error  // wrapped cause, compatible with errors.Is/As
}
```

See the [Error Handling guide](/xlfill/guides/error-handling/) for usage patterns.

### Progress and mode types

```go
type FillProgress struct {
    ProcessedRows int
    TotalRows     int    // estimated, may be 0
    CurrentSheet  string
    Elapsed       time.Duration
}

type Mode int
const (
    ModeSequential Mode = iota
    ModeStreaming
    ModeParallel
)

type ModeSuggestion struct {
    Mode        Mode
    Parallelism int      // only set when Mode == ModeParallel
    Reasons     []string // human-readable explanation
}
```

### CompiledTemplate

```go
type CompiledTemplate struct { /* ... */ }

func (ct *CompiledTemplate) Fill(data map[string]any, outputPath string) error
func (ct *CompiledTemplate) FillBytes(data map[string]any) ([]byte, error)
func (ct *CompiledTemplate) FillWriter(data map[string]any, w io.Writer) error
func (ct *CompiledTemplate) FillBatch(datasets []map[string]any, outputDir string, nameFunc func(i int, data map[string]any) string) ([]string, error)
```

`FillBatch` generates multiple files from the same compiled template. The `nameFunc` returns the filename for each dataset. Returns the list of generated file paths.

```go
compiled, _ := xlfill.Compile("template.xlsx")
files, err := compiled.FillBatch(datasets, "./reports/",
    func(i int, data map[string]any) string {
        return fmt.Sprintf("report_%s.xlsx", data["month"])
    },
)
// files: ["./reports/report_jan.xlsx", "./reports/report_feb.xlsx", ...]
```

### CellRef, AreaRef, Size

```go
type CellRef struct {
    Sheet string
    Row   int  // 0-based row index
    Col   int  // 0-based column index
}

type AreaRef struct {
    First CellRef
    Last  CellRef
}

type Size struct {
    Width  int
    Height int
}
```

### Command interface

Implement this to create [custom commands](/xlfill/guides/custom-commands/):

```go
type Command interface {
    Name() string
    ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error)
    Reset()
}
```

### Listener interfaces

```go
// Called before/after each cell transformation
type AreaListener interface {
    BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool
    AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer)
}

// Called after cell transformation for conditional styling
type StyleListener interface {
    StyleCell(target CellRef, value any, ctx *Context) *StyleOverride
}

type StyleOverride struct {
    Bold      *bool
    Italic    *bool
    FontColor *string  // hex color e.g. "#FF0000"
    FillColor *string  // hex background color
    FontSize  *float64
}
```

A listener can implement both `AreaListener` and `StyleListener`. See the [Area Listeners guide](/xlfill/guides/area-listeners/).

**Thread safety note:** When used with `WithParallelism`, listeners must be safe for concurrent use (use atomics or mutexes for shared state).

### Transformer interfaces

The `Transformer` is composed of three sub-interfaces for cleaner separation of concerns:

```go
type CellReader interface {
    GetCellData(ref CellRef) *CellData
    GetCommentedCells() []*CellData
    GetFormulaCells() []*CellData
}

type CellWriter interface {
    Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error
    ClearCell(ref CellRef) error
    SetFormula(ref CellRef, formula string) error
    SetCellValue(ref CellRef, value any) error
}

type SheetManager interface {
    GetSheetNames() []string
    DeleteSheet(name string) error
    CopySheet(src, dst string) error
    // ... and more
}

type Transformer interface {
    CellReader
    CellWriter
    SheetManager
    // ... plus target tracking, I/O, images, merges, hyperlinks
}
```

Custom Transformer implementations only need to satisfy the methods they use. The split makes it easier to reason about what each part does.

### ValidationIssue

```go
type Severity int
const (
    SeverityError   Severity = iota
    SeverityWarning
)

type ValidationIssue struct {
    Severity Severity
    CellRef  CellRef
    Message  string
}
```

`String()` formats as `[ERROR] Sheet1!A2: message` or `[WARN] Sheet1!A2: message`.

### Warning

```go
type Warning struct {
    Cell    CellRef
    Message string
}
```

Collected via `Filler.Warnings()` after processing.

### RowScanner interface

```go
type RowScanner interface {
    Next() bool
    Scan() (map[string]any, error)
    Close() error
}
```

Implement `RowScanner` for lazy data loading — rows are scanned one at a time during template processing instead of loading everything into memory upfront. Use with database cursors or CSV readers for large datasets.

### DeferredAction

```go
type DeferredAction struct {
    Command string
    Area    AreaRef
    Attrs   map[string]string
}
```

Represents a command that executes after all rows are written. Commands like `jx:table`, `jx:chart`, `jx:conditionalFormat`, `jx:group`, `jx:definedName`, and `jx:sparkline` use deferred execution to ensure correct output ranges.

## Data helpers

Convert common Go data sources to the `map[string]any` format XLFill expects:

### StructSliceToData

```go
func StructSliceToData(key string, slice any) map[string]any
```

Convert a slice of structs to a data map. The struct fields become map keys.

```go
type Employee struct {
    Name       string
    Department string
    Salary     float64
}

employees := []Employee{{Name: "Alice"}, {Name: "Bob"}}
data := xlfill.StructSliceToData("employees", employees)
// data = map[string]any{"employees": [...]any{map[Name:Alice ...], map[Name:Bob ...]}}
```

### JSONToData

```go
func JSONToData(jsonBytes []byte) (map[string]any, error)
```

Parse JSON bytes directly into a data map.

```go
data, err := xlfill.JSONToData([]byte(`{"employees": [{"Name": "Alice"}]}`))
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

### SQLRowsToData

```go
func SQLRowsToData(key string, rows *sql.Rows) (map[string]any, error)
```

Convert `*sql.Rows` from a database query into a data map. Column names become field names.

```go
rows, _ := db.Query("SELECT name, department, salary FROM employees")
data, err := xlfill.SQLRowsToData("employees", rows)
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

## Built-in functions

XLFill includes 16 built-in functions available in all `${...}` expressions. See the full [Built-in Functions guide](/xlfill/guides/built-in-functions/) for examples.

| Function | Description |
|----------|-------------|
| `hyperlink(url, display)` | Create a clickable hyperlink |
| `comment(text)` | Add a cell comment/note |
| `upper(s)` | Convert to uppercase |
| `lower(s)` | Convert to lowercase |
| `title(s)` | Convert to title case |
| `join(sep, items)` | Join a slice into a string |
| `formatNumber(value, format)` | Format a number |
| `formatDate(value, layout)` | Format a date/time |
| `coalesce(values...)` | First non-nil/non-empty value |
| `ifEmpty(value, fallback)` | Fallback for empty values |
| `sumBy(items, field)` | Sum a numeric field |
| `avgBy(items, field)` | Average a numeric field |
| `countBy(items, field)` | Count non-nil field values |
| `minBy(items, field)` | Minimum of a numeric field |
| `maxBy(items, field)` | Maximum of a numeric field |
| `t(key)` | Translate using i18n map |

Register custom functions with `WithFunction`:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithFunction("currency", func(args ...any) (any, error) {
        return fmt.Sprintf("$%.2f", args[0].(float64)), nil
    }),
)
```

## Data input

The `data` parameter accepts `map[string]any`. Values can be:

| Type | Example | Template access |
|------|---------|----------------|
| Primitives | `string`, `int`, `float64`, `bool` | `${name}`, `${count}` |
| Maps | `map[string]any` | `${employee.Name}` |
| Slices | `[]any`, `[]Employee` | Used in `jx:each(items="...")` |
| Structs | Any Go struct | `${emp.Name}` — fields by name |
| Byte slices | `[]byte` | Used in `jx:image(src="...")` |

Nested access works via dot notation: `${employee.Address.City}`.

## What's next?

For error handling patterns and data validation:

**[Error Handling &rarr;](/xlfill/guides/error-handling/)**

For streaming, parallel, and compiled template optimization:

**[Performance Tuning &rarr;](/xlfill/guides/performance-tuning/)**

Having trouble with a template? See the debugging toolkit:

**[Debugging & Troubleshooting &rarr;](/xlfill/guides/debugging/)**
