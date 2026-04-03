# xlfill 24-Feature Enhancement Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement all 24 features from the xlfill Enhancement Roadmap — built-in functions, 8 new commands, deferred post-processing, custom functions API, data helpers, batch mode, HTTP handler, and more — with full tests, zero regressions, race-detector clean, and comprehensive docs.

**Architecture:** All features hook into xlfill's existing extension points: CommandRegistry for new commands, Context.ToMap() for built-in functions, Options for new configuration, and a NEW deferred commands phase (inserted after area processing, before pre-write callback) for features that need final output ranges. New commands that need post-processing (table, conditionalFormat, group, chart, definedName, dataValidation preservation) register DeferredAction callbacks during ApplyAt, which the Filler executes after all areas complete.

**Tech Stack:** Go 1.25+, excelize/v2 (AddTable, AddChart, SetConditionalFormat, GroupRows, AddDataValidation, AddSparkline, ProtectSheet, SetPanes, InsertPageBreak, SetDefinedName, AddComment, SetDocProps), expr-lang/expr

---

## File Structure

### New source files to create:
| File | Responsibility |
|------|---------------|
| `builtins.go` | All built-in template functions (formatDate, coalesce, sumBy, ifEmpty, join, upper, lower, formatNumber, countBy, avgBy, minBy, maxBy, i18n, comment) |
| `deferred.go` | DeferredAction interface, DeferredRegistry, execution in Filler pipeline |
| `datavalidation.go` | jx:dataValidation command |
| `table.go` | jx:table command (deferred) |
| `conditionalformat.go` | jx:conditionalFormat command (deferred) |
| `group.go` | jx:group command (deferred) |
| `pagebreak.go` | jx:pageBreak command |
| `chart.go` | jx:chart command (deferred) |
| `autocolwidth.go` | jx:autoColWidth command |
| `freezepanes.go` | jx:freezePanes command |
| `definedname.go` | jx:definedName command (deferred) |
| `sparkline.go` | jx:sparkline command (deferred) |
| `protect.go` | jx:protect command |
| `include.go` | jx:include command |
| `helpers.go` | Data source helpers (StructSliceToData, JSONToData, SQLRowsToData) |
| `httphandler.go` | HTTPHandler helper |

### New test files:
| File | Tests |
|------|-------|
| `builtins_test.go` | All built-in functions |
| `deferred_test.go` | Deferred command pipeline |
| `newcommands_test.go` | All new commands (dataValidation, table, conditionalFormat, group, pageBreak, chart, autoColWidth, freezePanes, definedName, sparkline, protect, include) |
| `helpers_test.go` | Data source helpers, HTTPHandler, FillBatch |

### Existing files to modify:
| File | Changes |
|------|---------|
| `context.go` | Register built-in functions in ToMap(), add WithFunction support |
| `options.go` | Add WithFunction, WithI18n, WithDocumentProperties, selective streaming options |
| `command.go` | Register all new commands in NewCommandRegistry |
| `filler.go` | Add deferred commands phase, propagate new options |
| `xlfill.go` | Add FillWriter deferred phase, add HTTPHandler, FillBatch APIs |
| `compiled.go` | Propagate new options, add FillBatch method |
| `excelize_tx.go` | Add data validation preservation during Transform |
| `streaming.go` | Add selective sheet streaming |
| `automode.go` | Update template analysis for new commands |
| `describe.go` | Add describe support for all new commands |
| `validate.go` | Add validation for new command attributes |
| `each.go` | Data validation preservation during iteration |

---

## Chunk 1: Foundation — Deferred Commands Phase + Built-in Functions + Custom Functions API

This chunk is the prerequisite for 6+ features. It adds the deferred execution pipeline and the most requested feature (built-in functions).

### Task 1: Deferred Commands Architecture

**Files:**
- Create: `deferred.go`
- Modify: `filler.go` (add deferred phase after line 151)
- Modify: `xlfill.go` (wire deferred phase into FillWriter)
- Test: `deferred_test.go`

The DeferredAction pattern: commands register callbacks during ApplyAt that execute after ALL areas are processed — when final output ranges are known.

- [ ] **Step 1: Create deferred.go with DeferredAction interface and registry**

```go
// deferred.go
package xlfill

// DeferredAction represents a post-processing action that runs after all areas
// are processed but before the file is written. Used by commands that need to
// know final output ranges (jx:table, jx:conditionalFormat, jx:group, etc.).
type DeferredAction struct {
    Name     string
    Sheet    string
    StartRow int
    StartCol int
    EndRow   int  // -1 means "use actual expanded end"
    EndCol   int
    Execute  func(tx *ExcelizeTransformer, startRow, startCol, endRow, endCol int) error
}

// DeferredRegistry collects deferred actions during area processing.
type DeferredRegistry struct {
    actions []DeferredAction
}

func NewDeferredRegistry() *DeferredRegistry {
    return &DeferredRegistry{}
}

func (dr *DeferredRegistry) Add(action DeferredAction) {
    dr.actions = append(dr.actions, action)
}

func (dr *DeferredRegistry) Actions() []DeferredAction {
    return dr.actions
}

func (dr *DeferredRegistry) Reset() {
    dr.actions = dr.actions[:0]
}
```

- [ ] **Step 2: Wire deferred registry into Filler and Context**

Add `deferred *DeferredRegistry` field to Options. Pass it through Context so commands can register deferred actions during ApplyAt. Execute all deferred actions in FillWriter after area processing.

- [ ] **Step 3: Write test for deferred action pipeline**

```go
func TestDeferredActionExecution(t *testing.T) {
    // Create template, register a deferred action via a custom command,
    // verify it executes after area processing with correct output range
}
```

- [ ] **Step 4: Run tests, verify pass**
- [ ] **Step 5: Commit**

### Task 2: Built-in Template Functions (#1 + #6)

**Files:**
- Create: `builtins.go`
- Modify: `context.go` (register functions in ToMap)
- Test: `builtins_test.go`

- [ ] **Step 1: Create builtins.go with all functions**

Functions to implement:
- `formatDate(date, layout)` — Go-style date formatting
- `formatNumber(val, format)` — number formatting (#,##0.00)
- `upper(s)`, `lower(s)`, `title(s)` — string case
- `coalesce(values...)` — first non-nil
- `ifEmpty(val, default)` — default for empty/nil
- `join(items, sep)` — join slice to string
- `sumBy(items, field)` — aggregate sum
- `avgBy(items, field)` — aggregate average
- `countBy(items, field, value)` — count matching
- `minBy(items, field)` — aggregate minimum
- `maxBy(items, field)` — aggregate maximum

- [ ] **Step 2: Register in Context.ToMap()**
- [ ] **Step 3: Write comprehensive tests for each function**
- [ ] **Step 4: Run tests, verify all pass**
- [ ] **Step 5: Commit**

### Task 3: Custom Expression Functions API (#10/#14)

**Files:**
- Modify: `options.go` (add WithFunction option)
- Modify: `context.go` (merge custom functions into ToMap)
- Test: `builtins_test.go` (add custom function tests)

- [ ] **Step 1: Add WithFunction(name, fn) option**

```go
// options.go
customFunctions map[string]any

func WithFunction(name string, fn any) Option {
    return func(o *Options) {
        if o.customFunctions == nil {
            o.customFunctions = make(map[string]any)
        }
        o.customFunctions[name] = fn
    }
}
```

- [ ] **Step 2: Merge custom functions into Context**
- [ ] **Step 3: Write tests for WithFunction**
- [ ] **Step 4: Run tests, verify pass**
- [ ] **Step 5: Commit**

### Task 4: i18n / Resource Bundles (#13)

**Files:**
- Modify: `options.go` (add WithI18n)
- Modify: `builtins.go` (add i18n function)
- Test: `builtins_test.go`

- [ ] **Step 1: Add WithI18n(bundle map[string]string) option**
- [ ] **Step 2: Register i18n() function in builtins**
- [ ] **Step 3: Write tests**
- [ ] **Step 4: Run tests, verify pass**
- [ ] **Step 5: Commit**

### Task 5: comment() Built-in Function (#22)

**Files:**
- Modify: `builtins.go` (add CommentValue type and comment function)
- Modify: `excelize_tx.go` (handle CommentValue in Transform)
- Test: `builtins_test.go`

- [ ] **Step 1: Add CommentValue type and comment() function**
- [ ] **Step 2: Handle in ExcelizeTransformer.Transform**
- [ ] **Step 3: Write tests**
- [ ] **Step 4: Run all tests including regression**
- [ ] **Step 5: Commit**

---

## Chunk 2: Deferred Commands — table, conditionalFormat, group, chart, definedName

All these commands use the deferred phase from Chunk 1.

### Task 6: jx:dataValidation Command (#2)

**Files:**
- Create: `datavalidation.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement DataValidationCommand**

Attributes: type (list/integer/decimal/date/custom), source/formula, min/max, allowBlank, showError, errorTitle, errorMessage, lastCell

Uses excelize AddDataValidation in a deferred action (applies after jx:each expansion).

- [ ] **Step 2: Register in CommandRegistry**
- [ ] **Step 3: Write tests (dropdown list, integer range, within each loop)**
- [ ] **Step 4: Run tests**
- [ ] **Step 5: Commit**

### Task 7: jx:table Command (#3)

**Files:**
- Create: `table.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement TableCommand (deferred)**

Attributes: name, style (e.g. TableStyleMedium9), showTotals, showFirstColumn, showLastColumn, lastCell

Deferred action calls excelize AddTable after area expansion to get final range.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 8: jx:conditionalFormat Command (#4)

**Files:**
- Create: `conditionalformat.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement ConditionalFormatCommand (deferred)**

Attributes: type (colorScale/dataBar/iconSet/cellIs), operator, value, format, minColor/maxColor, color, style, lastCell

Deferred action calls excelize SetConditionalFormat.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 9: jx:group Command (#5)

**Files:**
- Create: `group.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement GroupCommand (deferred)**

Attributes: collapsed (bool), level (int), lastCell

Deferred action calls excelize GroupRows/GroupColumns.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 10: jx:chart Command (#8)

**Files:**
- Create: `chart.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement ChartCommand (deferred)**

Attributes: type (bar/line/pie/scatter/area), title, series (data range expression), categories, position, width, height, lastCell

Deferred action calls excelize AddChart after expansion.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 11: jx:definedName Command (#12)

**Files:**
- Create: `definedname.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement DefinedNameCommand (deferred)**

Attributes: name, scope (sheet/workbook), lastCell

Deferred action calls excelize SetDefinedName with expanded range.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 12: jx:sparkline Command (#19)

**Files:**
- Create: `sparkline.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement SparklineCommand (deferred)**

Attributes: type (line/column/winLoss), data (range expression), lastCell

Deferred action calls excelize AddSparkline.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

---

## Chunk 3: Simple Commands — pageBreak, autoColWidth, freezePanes, protect, include

These commands don't need deferred execution — they apply immediately.

### Task 13: jx:pageBreak Command (#7)

**Files:**
- Create: `pagebreak.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement PageBreakCommand**

Attributes: lastCell. Calls excelize InsertPageBreak at target row.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 14: jx:autoColWidth Command (#9)

**Files:**
- Create: `autocolwidth.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement AutoColWidthCommand**

Attributes: lastCell. After area processing, estimates column width from content length and calls excelize SetColWidth.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 15: jx:freezePanes Command (#11)

**Files:**
- Create: `freezepanes.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement FreezePanesCommand**

Attributes: row (freeze below this row), col (freeze right of this col), lastCell

Calls excelize SetPanes.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 16: jx:protect Command (#20)

**Files:**
- Create: `protect.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement ProtectCommand**

Attributes: password, allowSelectUnlockedCells, allowSelectLockedCells, allowSort, allowFilter, lastCell

Calls excelize ProtectSheet.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

### Task 17: jx:include Command (#23)

**Files:**
- Create: `include.go`
- Modify: `command.go` (register)
- Test: `newcommands_test.go`

- [ ] **Step 1: Implement IncludeCommand**

Attributes: template (file path or context key), sheet (source sheet name), area (source range e.g. "A1:F3"), lastCell

Opens the included template, reads the specified area, copies cells to the target position.

- [ ] **Step 2: Register and test**
- [ ] **Step 3: Commit**

---

## Chunk 4: Data Validation Preservation + Remaining Features

### Task 18: Data Validation Preservation During jx:each (#15)

**Files:**
- Modify: `excelize_tx.go` (preserve validation rules during Transform)
- Modify: `each.go` (copy validation to expanded rows)
- Test: `newcommands_test.go`

- [ ] **Step 1: Read validation rules from template cells during readAllCellData**
- [ ] **Step 2: Copy validation rules to target cells during Transform**
- [ ] **Step 3: Write tests (dropdown preserved across 5 loop iterations)**
- [ ] **Step 4: Run tests**
- [ ] **Step 5: Commit**

### Task 19: WithDocumentProperties Option (#21)

**Files:**
- Modify: `options.go` (add WithDocumentProperties)
- Modify: `xlfill.go` (apply in FillWriter before write)
- Test: `newcommands_test.go`

- [ ] **Step 1: Add option and apply via excelize SetDocProps**
- [ ] **Step 2: Write tests**
- [ ] **Step 3: Commit**

### Task 20: CompiledTemplate.FillBatch (#16)

**Files:**
- Modify: `compiled.go` (add FillBatch method)
- Test: `helpers_test.go`

- [ ] **Step 1: Add FillBatch method**

```go
func (ct *CompiledTemplate) FillBatch(
    items []map[string]any,
    nameFn func(index int, data map[string]any) string,
) error
```

Generates one output file per item using the compiled template.

- [ ] **Step 2: Write tests (3 items → 3 files)**
- [ ] **Step 3: Commit**

### Task 21: HTTPHandler Helper (#17)

**Files:**
- Create: `httphandler.go`
- Test: `helpers_test.go`

- [ ] **Step 1: Implement HTTPHandler**

```go
func HTTPHandler(compiled *CompiledTemplate, dataFn func(*http.Request) (map[string]any, string, error)) http.HandlerFunc
```

Sets Content-Type, Content-Disposition, streams the filled template.

- [ ] **Step 2: Write tests using httptest**
- [ ] **Step 3: Commit**

### Task 22: Data Source Helpers (#18)

**Files:**
- Create: `helpers.go`
- Test: `helpers_test.go`

- [ ] **Step 1: Implement helpers**

```go
func StructSliceToData[T any](items []T) []map[string]any
func JSONToData(r io.Reader) (map[string]any, error)
func SQLRowsToData(rows RowScanner) ([]map[string]any, error)

type RowScanner interface {
    Columns() ([]string, error)
    Next() bool
    Scan(dest ...any) error
}
```

- [ ] **Step 2: Write tests**
- [ ] **Step 3: Commit**

### Task 23: Selective Sheet Streaming (#24)

**Files:**
- Modify: `options.go` (add WithStreamingSheets)
- Modify: `streaming.go` (support per-sheet streaming)
- Modify: `xlfill.go` (create StreamingTransformer per sheet)
- Test: `streaming_parallel_test.go`

- [ ] **Step 1: Add WithStreamingSheets(sheets ...string) option**
- [ ] **Step 2: Create StreamingTransformer only for specified sheets**
- [ ] **Step 3: Write tests**
- [ ] **Step 4: Run tests**
- [ ] **Step 5: Commit**

---

## Chunk 5: Integration, Testing, Review

### Task 24: Update automode.go for New Commands

**Files:**
- Modify: `automode.go` (detect new commands in template analysis)
- Modify: `describe.go` (describe all new commands)
- Modify: `validate.go` (validate new command attributes)

- [ ] **Step 1: Update analyzeArea for all new command types**
- [ ] **Step 2: Update describeCommandAttrs for all new command types**
- [ ] **Step 3: Update validateCommandAttributes for new commands**
- [ ] **Step 4: Run tests**
- [ ] **Step 5: Commit**

### Task 25: Full Regression Test Suite

- [ ] **Step 1: Run all tests**: `go test ./... -count=1`
- [ ] **Step 2: Run race detector**: `go test ./... -count=1 -race`
- [ ] **Step 3: Run go vet**: `go vet ./...`
- [ ] **Step 4: Run benchmarks**: `go test -bench=. -benchmem -count=1`
- [ ] **Step 5: Compare benchmark numbers to baseline (should not degrade)**

### Task 26: Code Review Iteration 1

- [ ] **Step 1: Review all new source files for race conditions, memory leaks, panics**
- [ ] **Step 2: Review all test files for correctness, flakiness, cleanup**
- [ ] **Step 3: Fix all issues found**
- [ ] **Step 4: Re-run full test suite with race detector**

### Task 27: Code Review Iteration 2

- [ ] **Step 1: Review all modified files for API consistency and error handling**
- [ ] **Step 2: Verify all deferred commands work with streaming and parallel modes**
- [ ] **Step 3: Fix any issues**
- [ ] **Step 4: Final regression check**

---

## Chunk 6: Documentation Overhaul + Push

### Task 28: Documentation Overhaul

All docs under `docs/src/content/docs/`. Create/update:

- [ ] **Step 1: Create command pages** for each new command (dataValidation, table, conditionalFormat, group, pageBreak, chart, autoColWidth, freezePanes, definedName, sparkline, protect, include)
- [ ] **Step 2: Update commands-overview.md** with all 20 commands
- [ ] **Step 3: Update reference/api.md** with all new functions, options, types
- [ ] **Step 4: Create guides/built-in-functions.md** with all functions + examples
- [ ] **Step 5: Update guides/performance-tuning.md** with deferred commands info
- [ ] **Step 6: Create reference/comparison.md** — feature comparison vs JXLS, EPPlus, XlsxWriter, Aspose
- [ ] **Step 7: Update index.mdx** landing page with new feature highlights
- [ ] **Step 8: Update README.md** with complete feature list
- [ ] **Step 9: Build docs**: `cd docs && npx astro build`
- [ ] **Step 10: Verify build success (0 errors)**

### Task 29: Commit and Push

- [ ] **Step 1: Stage all files**
- [ ] **Step 2: Create commit**
- [ ] **Step 3: Push to origin/main**
- [ ] **Step 4: Verify GitHub Actions workflow triggered**
