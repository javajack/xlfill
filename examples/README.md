# xlfill Example

A working example that creates a template, fills it with employee data, and prints the output.

## Run

```bash
cd examples
go run .
```

## What It Does

1. Creates a template `.xlsx` with `jx:area` and `jx:each` commands in cell comments
2. Fills the template with a list of employees
3. Writes `output.xlsx`
4. Reads back and prints the result to verify
