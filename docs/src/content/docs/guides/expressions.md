---
title: Expressions
description: Everything you can put inside ${...} in XLFill templates.
---

Expressions are the values inside `${...}` in your template cells. They're evaluated against the current data context and replaced with the result.

## Field access

The most common use — access fields on your data:

```
${employee.Name}
${dept.Manager.Email}
${company.Address.City}
```

This works with Go structs, maps, or any nested combination.

## Arithmetic

Do math directly in the template:

```
${price * quantity}
${subtotal + tax}
${total / count}
${score * 100 / maxScore}
```

## Comparisons and logic

```
${age >= 18}              // true or false
${status == "active"}     // string comparison
${price > 100 && inStock} // logical AND
${!expired}               // negation
```

## Ternary expressions

Conditional values without needing a `jx:if` command:

```
${age >= 18 ? "Adult" : "Minor"}
${score > 90 ? "A" : score > 80 ? "B" : "C"}
${active ? "Yes" : "No"}
```

Great for status columns, pass/fail indicators, and conditional labels.

## Indexing

Access items by position:

```
${items[0].Name}       // first item
${matrix[row][col]}    // 2D access
${months[quarterStart]}  // dynamic index
```

## String concatenation

```
${firstName + " " + lastName}
${prefix + phoneNumber}
```

## Mixed text and expressions

A single cell can contain plain text alongside expressions:

```
Employee: ${e.Name} (${e.Department})
Total: ${amount} USD
Report generated for ${company.Name}
```

XLFill replaces only the `${...}` parts and keeps the surrounding text.

## Built-in variables

These are available in every expression automatically:

| Variable | Type | Description |
|----------|------|-------------|
| `_row` | int | Current output row number (1-based) |
| `_col` | int | Current output column index (0-based) |

```
Row ${_row}: ${e.Name}
```

Useful for row numbering, conditional formatting logic, or debugging.

## Built-in functions

### hyperlink(url, display)

Creates a clickable Excel hyperlink:

```
${hyperlink("https://example.com", "Click here")}
${hyperlink(e.ProfileURL, e.Name)}
```

The cell becomes a real Excel hyperlink — blue, underlined, clickable.

## Custom delimiters

If `${...}` conflicts with content in your spreadsheet (rare, but possible), change the delimiters:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithExpressionNotation("<<", ">>"),
)
```

Then use `<<e.Name>>` in your template instead.

## Expression engine

Under the hood, expressions are powered by [expr-lang/expr](https://github.com/expr-lang/expr), a fast and safe expression evaluator for Go. It supports a rich syntax including:

- All Go operators
- String functions
- Array/slice functions
- Type coercion
- And more — see the [expr-lang documentation](https://expr-lang.org/docs/language-definition) for the full reference

Expressions are compiled once and cached, so even templates with thousands of rows evaluate quickly (~5 million evaluations per second).

## What's next?

Expressions fill individual cells with values. But to control **structure** — loops, conditions, grids — you need commands. Let's see how they work.

**[Commands Overview &rarr;](/guides/commands-overview/)**
