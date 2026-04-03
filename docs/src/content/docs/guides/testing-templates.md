---
title: "Testing Excel Templates in CI/CD"
description: "Validate templates at build time, check data contracts, enforce strict mode, and test generated output — all without opening Excel."
---

Templates live outside your Go code. That's their strength — but it also means a typo in `${e.Nmae}` won't show up until runtime. Unless you test for it.

XLFill gives you four layers of template testing, all runnable from `go test` and CI pipelines. Catch errors at build time, not in production.

## Layer 1: Validate() — syntax checking without data

`Validate` parses a template and checks for structural errors *without needing any data*:

```go
func TestTemplateSyntax(t *testing.T) {
    issues, err := xlfill.Validate("templates/report.xlsx")
    if err != nil {
        t.Fatalf("Cannot open template: %v", err)
    }

    for _, issue := range issues {
        t.Errorf("Template issue: %s", issue)
    }
}
```

What it catches:
- **Invalid expression syntax** — `${e.Name +}` (trailing operator), `${e.Name` (unclosed brace)
- **Invalid command attributes** — `jx:each(item="employees")` (should be `items`, not `item`)
- **Mismatched areas** — `lastCell` outside the parent area
- **Unknown commands** — `jx:loop(...)` (no such command)

This is your first line of defense. Run it on every template in your project:

```go
func TestAllTemplates(t *testing.T) {
    templates, _ := filepath.Glob("templates/*.xlsx")
    for _, tmpl := range templates {
        t.Run(filepath.Base(tmpl), func(t *testing.T) {
            issues, err := xlfill.Validate(tmpl)
            if err != nil {
                t.Fatalf("Cannot parse: %v", err)
            }
            for _, issue := range issues {
                t.Errorf("%s", issue)
            }
        })
    }
}
```

> **Pro tip:** Add this test to your CI pipeline. It runs in milliseconds and catches every template typo before deployment.

## Layer 2: ValidateData() — data contract testing

`ValidateData` goes deeper. It checks that your data map provides every variable the template expects:

```go
func TestReportDataContract(t *testing.T) {
    data := map[string]any{
        "employees": []map[string]any{
            {"Name": "Alice", "Department": "Engineering", "Salary": 95000},
        },
    }

    issues, err := xlfill.ValidateData("templates/report.xlsx", data)
    if err != nil {
        t.Fatalf("Cannot parse template: %v", err)
    }

    for _, issue := range issues {
        t.Errorf("Data contract violation: %s", issue)
    }
}
```

What it catches:
- **Missing top-level variables** — template uses `${companyName}` but data doesn't have it
- **Missing nested fields** — template uses `${e.Email}` but Employee struct has no `Email` field
- **Type mismatches** — `jx:each(items="employees")` but `employees` is a string, not a slice

This is particularly valuable when your data structure changes. If a developer renames `Salary` to `AnnualPay` in a struct, the test fails immediately.

```go
// Test that the API response matches template expectations
func TestAPIResponseMatchesTemplate(t *testing.T) {
    // Use real API response structure (sample data)
    resp := fetchSampleAPIResponse()
    data := xlfill.JSONToData(resp)

    issues, _ := xlfill.ValidateData("templates/api_report.xlsx", data)
    for _, issue := range issues {
        t.Errorf("%s", issue)
    }
}
```

## Layer 3: WithStrictMode — fail on anything suspicious

Strict mode turns warnings into errors. In normal mode, an unknown command like `jx:loop(...)` produces a warning. In strict mode, it produces an error:

```go
func TestStrictMode(t *testing.T) {
    data := buildTestData()

    err := xlfill.Fill("templates/report.xlsx", "test_output.xlsx", data,
        xlfill.WithStrictMode(true),
    )
    if err != nil {
        t.Fatalf("Strict mode failure: %v", err)
    }

    os.Remove("test_output.xlsx") // clean up
}
```

Strict mode catches:
- Unknown commands (typos like `jx:eaach`)
- Unknown attributes (typos like `ietms` instead of `items`)
- Expressions that evaluate to nil when a value was expected

> **Gotcha:** Don't enable strict mode in production unless you've tested thoroughly. It's meant for CI — to catch things during development that you'd otherwise miss.

## Layer 4: Describe() — snapshot testing

`Describe` returns a human-readable tree of your template's structure. Use it for snapshot testing — verify that the template hasn't changed unexpectedly:

```go
func TestTemplateStructure(t *testing.T) {
    description, err := xlfill.Describe("templates/report.xlsx")
    if err != nil {
        t.Fatal(err)
    }

    expected := `Template: templates/report.xlsx
Sheet1!A1:C4 area (3x4)
  Commands:
    Sheet1!A2 each (3x3) items="departments" var="dept"
      Sheet1!A2:C4 area (3x3)
        Sheet1!A4 each (3x1) items="dept.Employees" var="e"
          Sheet1!A4:C4 area (3x1)
            Expressions:
              A4: ${e.Name}
              B4: ${e.Role}
              C4: ${e.Salary}
`

    if description != expected {
        t.Errorf("Template structure changed:\n\nGot:\n%s\nExpected:\n%s", description, expected)
    }
}
```

If someone modifies the template (adds a column, changes an expression, moves a command), this test catches it. It's like a snapshot test for your template's logical structure.

> **Pro tip:** Store the expected description in a testdata file for easier maintenance:
> ```go
> expected, _ := os.ReadFile("testdata/report_structure.txt")
> if description != string(expected) { ... }
> ```

## Testing generated output

For integration tests, generate the output and inspect it programmatically with excelize:

```go
func TestGeneratedOutput(t *testing.T) {
    data := map[string]any{
        "employees": []map[string]any{
            {"Name": "Alice", "Department": "Engineering", "Salary": 95000},
            {"Name": "Bob", "Department": "Sales", "Salary": 72000},
        },
    }

    outputPath := filepath.Join(t.TempDir(), "output.xlsx")
    err := xlfill.Fill("templates/report.xlsx", outputPath, data)
    if err != nil {
        t.Fatalf("Fill failed: %v", err)
    }

    // Open output with excelize and verify contents
    f, err := excelize.OpenFile(outputPath)
    if err != nil {
        t.Fatalf("Cannot open output: %v", err)
    }
    defer f.Close()

    // Check cell values
    name, _ := f.GetCellValue("Sheet1", "A2")
    if name != "Alice" {
        t.Errorf("Expected 'Alice' in A2, got '%s'", name)
    }

    // Check row count (header + 2 data rows)
    rows, _ := f.GetRows("Sheet1")
    if len(rows) != 3 {
        t.Errorf("Expected 3 rows, got %d", len(rows))
    }

    // Check that formatting was preserved
    style, _ := f.GetCellStyle("Sheet1", "A1")
    font, _ := f.GetStyle(style)
    if font != nil && !font.Font.Bold {
        t.Error("Header should be bold")
    }
}
```

## Golden file pattern

For complex reports, use golden files — a known-good output that you compare against:

```go
func TestGoldenFile(t *testing.T) {
    data := loadFixtureData("testdata/report_fixture.json")

    outputPath := filepath.Join(t.TempDir(), "output.xlsx")
    xlfill.Fill("templates/report.xlsx", outputPath, data)

    // Compare output bytes with golden file
    got, _ := os.ReadFile(outputPath)
    golden, _ := os.ReadFile("testdata/golden/report.xlsx")

    if !bytes.Equal(got, golden) {
        // If updating golden files, set UPDATE_GOLDEN=1
        if os.Getenv("UPDATE_GOLDEN") == "1" {
            os.MkdirAll("testdata/golden", 0o755)
            os.WriteFile("testdata/golden/report.xlsx", got, 0o644)
            t.Log("Golden file updated")
            return
        }
        t.Error("Output differs from golden file. Run with UPDATE_GOLDEN=1 to update.")
    }
}
```

Update golden files when you intentionally change templates:
```bash
UPDATE_GOLDEN=1 go test ./...
```

> **Gotcha:** Excel files contain metadata (timestamps, creator info) that changes between runs. Use `WithDocumentProperties` with fixed values in tests to make output deterministic:
> ```go
> xlfill.Fill("template.xlsx", outputPath, data,
>     xlfill.WithDocumentProperties(xlfill.DocProperties{
>         Author: "Test", Created: fixedTime,
>     }),
> )
> ```

## Edge case testing

Test with data that breaks assumptions:

```go
func TestEdgeCases(t *testing.T) {
    cases := []struct {
        name string
        data map[string]any
    }{
        {
            name: "empty slice",
            data: map[string]any{"employees": []map[string]any{}},
        },
        {
            name: "nil values",
            data: map[string]any{
                "employees": []map[string]any{
                    {"Name": nil, "Salary": nil},
                },
            },
        },
        {
            name: "unicode",
            data: map[string]any{
                "employees": []map[string]any{
                    {"Name": "Müller", "Department": "日本語"},
                },
            },
        },
        {
            name: "very long text",
            data: map[string]any{
                "employees": []map[string]any{
                    {"Name": strings.Repeat("A", 32767)}, // Excel cell limit
                },
            },
        },
        {
            name: "special characters",
            data: map[string]any{
                "employees": []map[string]any{
                    {"Name": `O'Brien & "Friends" <test>`},
                },
            },
        },
    }

    for _, tc := range cases {
        t.Run(tc.name, func(t *testing.T) {
            outputPath := filepath.Join(t.TempDir(), "output.xlsx")
            err := xlfill.Fill("templates/report.xlsx", outputPath, tc.data)
            if err != nil {
                t.Errorf("Failed on %s: %v", tc.name, err)
            }
        })
    }
}
```

## CI pipeline integration

### GitHub Actions example

```yaml
name: Test Templates
on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-go@v5
        with:
          go-version: '1.24'
      - name: Test templates
        run: go test ./... -v -run TestTemplate
```

### Pre-commit hook

Add template validation to a pre-commit hook so broken templates never make it into the repo:

```bash
#!/bin/sh
# .git/hooks/pre-commit

# Check if any template files changed
TEMPLATES=$(git diff --cached --name-only --diff-filter=ACM | grep '\.xlsx$')
if [ -n "$TEMPLATES" ]; then
    echo "Validating Excel templates..."
    go test ./... -run TestAllTemplates -count=1
    if [ $? -ne 0 ]; then
        echo "Template validation failed. Fix the issues before committing."
        exit 1
    fi
fi
```

## Tips and tricks

- **Add `Validate` to pre-commit hooks.** It runs in milliseconds and catches typos before they hit CI.

- **Use `ValidateData` in integration tests.** When your data structures change (struct fields renamed, API response shape changes), this catches the mismatch instantly.

- **Test with empty slices.** The most common production bug is "template works with data but crashes with an empty list." Test this explicitly.

- **Use `t.TempDir()` for output files.** Go cleans them up automatically after the test.

- **Pin template metadata for deterministic output.** Use fixed `DocProperties` in tests to avoid timestamp-based diffs.

- **Run strict mode in CI, not production.** Strict mode surfaces issues you want to know about during development, but shouldn't crash your production server.

## What's next?

- Build a complete testing pipeline for batch reports: **[Batch Generation &rarr;](/xlfill/guides/batch-generation/)**
- Learn about error types and handling: **[Error Handling &rarr;](/xlfill/guides/error-handling/)**
- Validate templates before serving from HTTP: **[Excel Export APIs &rarr;](/xlfill/guides/excel-export-api/)**
