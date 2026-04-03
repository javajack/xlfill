---
title: "jx:dataValidation"
description: Add dropdown lists, integer/decimal constraints, and input messages to cells — all from your template.
---

`jx:dataValidation` adds Excel data validation rules to cells during template processing. Dropdowns, integer ranges, decimal constraints, date ranges, text length limits — configured right in the template, populated from your data.

## Syntax

```
jx:dataValidation(lastCell="B1" type="list" formula1="choices" allowBlank="true")
jx:dataValidation(lastCell="C1" type="whole" operator="between" formula1="1" formula2="100")
jx:dataValidation(lastCell="D1" type="decimal" operator="greaterThan" formula1="0")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the validation area |
| `type` | Yes | Validation type: `list`, `whole`, `decimal`, `date`, `textLength`, `custom` |
| `formula1` | Yes | Primary constraint — for `list`, an expression resolving to a comma-separated string or slice; for numeric types, the min/value |
| `formula2` | No | Secondary constraint — the max value for `between`/`notBetween` operators |
| `operator` | No | Comparison: `between`, `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual` |
| `allowBlank` | No | Allow empty cells (default: `true`) |
| `showInputMessage` | No | Show tooltip when cell is selected (default: `false`) |
| `inputTitle` | No | Input message title |
| `inputMessage` | No | Input message body |
| `showErrorMessage` | No | Show error popup on invalid input (default: `false`) |
| `errorTitle` | No | Error popup title |
| `errorMessage` | No | Error popup message |
| `errorStyle` | No | Error style: `stop`, `warning`, `information` (default: `stop`) |

## When to use jx:dataValidation

Use it whenever your generated spreadsheet will be **edited by humans** after generation. Data validation prevents bad input at the source:

- Dropdown lists for status fields, categories, or country codes
- Integer constraints on quantity columns
- Decimal ranges on price/discount fields
- Date ranges on scheduling sheets
- Text length limits on comment fields

## Example: Dropdown list

The most common use case — a dropdown list populated from your data:

```
Cell B2 comment:
  jx:dataValidation(lastCell="B2" type="list" formula1="statusOptions" showInputMessage="true" inputTitle="Status" inputMessage="Pick a status")
```

```go
data := map[string]any{
    "statusOptions": "Open,In Progress,Closed,Deferred",
}
```

When a user clicks cell B2 in the output, they get a dropdown with four choices and a helpful tooltip.

## Example: Integer range

Constrain a quantity field to whole numbers between 1 and 1000:

```
Cell C2 comment:
  jx:dataValidation(lastCell="C2" type="whole" operator="between" formula1="1" formula2="1000" showErrorMessage="true" errorTitle="Invalid quantity" errorMessage="Enter a whole number between 1 and 1000" errorStyle="stop")
```

No Go code needed — the constraints are literals in the template. If a user types `0` or `1500`, Excel blocks it with your custom error message.

## Example: Decimal with minimum

Ensure a price field is always positive:

```
Cell D2 comment:
  jx:dataValidation(lastCell="D2" type="decimal" operator="greaterThan" formula1="0" showErrorMessage="true" errorTitle="Invalid price" errorMessage="Price must be greater than zero")
```

## Inside loops

Combine with `jx:each` to apply validation to every generated row:

```
Cell A1 comment:
  jx:area(lastCell="D5")
  jx:each(items="items" var="item" lastCell="D1")

Cell B1 comment:
  jx:dataValidation(lastCell="B1" type="list" formula1="categories")
```

Every row gets the same dropdown. The validation rule is applied per-row during expansion.

## Deferred execution

`jx:dataValidation` uses deferred execution — the validation rules are collected during template processing and applied in a single batch after all rows are written. This is both faster and ensures correct cell references even when rows shift during expansion.

## What's next?

Add structured Excel tables with auto-filter and banded rows:

**[jx:table &rarr;](/xlfill/commands/table/)**
