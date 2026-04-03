---
title: Error Handling
description: Structured errors, warnings, data validation, and strict mode in XLFill.
---

XLFill provides a layered error system designed for production use. From quick checks during development to programmatic error routing in server-side pipelines — every error tells you *what* went wrong, *where*, and *what kind* of fix is needed.

## Structured error types

Every error returned by XLFill is (or wraps) an `XLFillError` with a `Kind` field that categorizes the problem:

```go
type XLFillError struct {
    Kind    ErrorKind  // Template, Data, or Runtime
    Cell    CellRef    // which cell, if applicable
    Command string     // which command, if applicable
    Message string
    Err     error      // wrapped cause
}
```

### Error kinds

| Kind | Meaning | Who fixes it |
|------|---------|-------------|
| `ErrTemplate` | Template structure problem — bad `jx:` syntax, missing area, invalid cell refs | Edit the `.xlsx` template |
| `ErrData` | Data/expression problem — missing variable, wrong type, evaluation failure | Fix the data map or expression |
| `ErrRuntime` | I/O or system problem — file not found, write failure, excelize error | Retry, fix permissions, check disk |

### Programmatic error handling

Use `errors.As` to route errors in your application:

```go
err := xlfill.Fill("template.xlsx", "report.xlsx", data)
if err != nil {
    var xlErr *xlfill.XLFillError
    if errors.As(err, &xlErr) {
        switch xlErr.Kind {
        case xlfill.ErrTemplate:
            log.Printf("Template bug at %s: %s", xlErr.Cell, xlErr.Message)
            // Alert the template author
        case xlfill.ErrData:
            log.Printf("Data issue: %s", xlErr.Message)
            // Return 400 to the API caller
        case xlfill.ErrRuntime:
            log.Printf("System error: %s", xlErr.Message)
            // Retry or escalate
        }
    }
}
```

This replaces string-matching on error messages — which is fragile — with a clean type switch.

## Warnings: catch problems that don't fail

Not every issue is fatal. XLFill collects **warnings** for problems that don't stop processing but might produce unexpected output.

The most common warning: **unknown commands with "did you mean?" suggestions**.

```go
filler := xlfill.NewFiller(xlfill.WithTemplate("template.xlsx"))
err := filler.Fill(data, "output.xlsx")

for _, w := range filler.Warnings() {
    fmt.Println(w)
}
// Output: [WARN] Sheet1!A5: unknown command "eache" (did you mean "each"?)
```

Previously, a typo like `jx:eache` was **silently ignored** — no error, no output, just confusion. Now it's caught and reported.

## Strict mode: warnings become errors

For CI pipelines and automated builds, use `WithStrictMode(true)` to turn all warnings into hard errors:

```go
err := xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithStrictMode(true),
)
// Error: [xlfill:template] Sheet1!A5: eache: unknown command "eache" (did you mean "each"?)
```

This is the recommended setting for production pipelines where silent issues are unacceptable.

## Data contract validation

`ValidateData` goes beyond syntax checking — it verifies that your data map actually satisfies the template's requirements:

```go
issues, err := xlfill.ValidateData("template.xlsx", data)
for _, issue := range issues {
    fmt.Println(issue)
}
// [WARN] Sheet1!A1: expression ${companyName} references variable "companyName" which is not in data
```

**What it checks:**
- Extracts all `${...}` variable references from cell values and formulas
- Cross-references against your data map keys
- Recognizes command-provided variables (`e` from `jx:each`, `i` from `jx:repeat`) so they're not flagged
- Filters out expr-lang built-in functions (`len`, `filter`, `true`, etc.)

**When to use it:**
- During development to verify template/data alignment
- In CI to catch regressions when data shapes change
- Before expensive processing to fail fast

```go
func TestReportDataContract(t *testing.T) {
    issues, err := xlfill.ValidateData("templates/monthly.xlsx", buildReportData())
    require.NoError(t, err)
    assert.Empty(t, issues, "data doesn't satisfy template: %v", issues)
}
```

## What's next?

For runtime debugging tools (trace mode, describe, validate):

**[Debugging & Troubleshooting &rarr;](/xlfill/guides/debugging/)**

For the complete options reference:

**[API Reference &rarr;](/xlfill/reference/api/)**
