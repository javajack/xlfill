# xlfill Enhancement Plan

## Baseline
- 505 tests passing, 95.3% coverage
- All enhancements must maintain full backward compatibility (zero regression)

## Results (Phase 1: Enhancements 1-8)
- **536 tests passing**, 94.0% coverage
- All 505 original tests pass (zero regression)
- 31 new tests added for enhancements
- 1 previously skipped test (multisheet) now passes

## Results (Phase 2: Enhancements 9-23)
- **616 tests passing**, 90.5% coverage
- All 536 Phase 1 tests pass (zero regression)
- 80 new tests added for enhancements 9-23
- Performance: 17% speedup on 10K-row benchmark (303ms → 250ms)

## Enhancement 1: Nested Commands
**Status:** Done
**Impact:** High — unlocks each-in-each, if-in-each, and all hierarchical templates
**What:** Reworked `BuildAreas` in `filler.go` to build a command tree. Commands whose area is strictly smaller than another command's area become children of that command. Equal-sized commands remain siblings.
**Files:** `filler.go` (nesting algorithm, `getCommandArea`, `sortAreaBindings`), `area.go` (`transformCell` helper)
**Tests:** Nested each (departments→employees), if-inside-each, three-level nesting, same-row different-scope

## Enhancement 2: Multisheet Each
**Status:** Done
**Impact:** High — `multisheet` attribute now fully functional
**What:** Added `applyMultiSheet` method to `EachCommand`. Copies template sheet per item, processes each copy, deletes the template sheet.
**Files:** `each.go` (`applyMultiSheet`, `toStringSlice`)
**Tests:** Dynamic sheet names from context, unskipped jxls3 parity test

## Enhancement 3: Recalculate Formulas on Open
**Status:** Done
**Impact:** High — one option, solves stale formula display
**What:** Added `WithRecalculateOnOpen(bool)` option. Uses excelize `SetCalcProps` with `FullCalcOnLoad`.
**Files:** `options.go`, `xlfill.go`, `excelize_tx.go`, `transformer.go`
**Tests:** Verify calc props set, verify default behavior

## Enhancement 4: Hyperlinks
**Status:** Done
**Impact:** Medium-High — common in business reports
**What:** Added `hyperlink(url, title)` built-in function. Returns `HyperlinkValue` that the transformer writes as display text + clickable link. Supports external URLs, mailto, and document locations.
**Files:** `hyperlink.go` (new), `context.go` (built-in function), `excelize_tx.go` (HyperlinkValue handling), `transformer.go` (`SetCellHyperLink`)
**Tests:** Basic hyperlink, hyperlinks in each loop, HyperlinkValue unit tests

## Enhancement 5: Area Listeners (Before/After Transform hooks)
**Status:** Done
**Impact:** Medium — enables conditional row/cell styling, logging, validation
**What:** Added `AreaListener` interface with `BeforeTransformCell`/`AfterTransformCell`. `WithAreaListener()` option. Listeners propagated to all areas including nested command areas.
**Files:** `listener.go` (new), `options.go`, `filler.go` (`propagateListeners`), `area.go` (`transformCell`)
**Tests:** Call count verification, skip-transform listener, listener with nested commands

## Enhancement 6: Parameterized Formulas
**Status:** Done
**Impact:** Medium — enables data-driven constants in formulas
**What:** Formula cells containing `${...}` expressions now have those expressions substituted before writing the formula. E.g., `=A1*${taxRate}` becomes `=A1*0.2`.
**Files:** `excelize_tx.go` (in `Transform` method)
**Tests:** Single variable substitution, multiple variables in one formula

## Enhancement 7: Auto Row Height
**Status:** Done
**Impact:** Medium — prevents text clipping on wrapped content
**What:** Added `jx:autoRowHeight` command. Processes its area then sets each output row to auto-height.
**Files:** `autorowheight.go` (new), `command.go` (registry), `filler.go` (`attachArea`, `propagateListeners`, `getCommandArea`)
**Tests:** Basic auto-height, auto-height with each loop

## Enhancement 8: Built-in Row/Col Context Variables
**Status:** Done
**Impact:** Medium — ergonomic improvement for row numbering and position-aware templates
**What:** `_row` (1-based output row) and `_col` (0-based output column) automatically injected into context during every cell transformation. Available in expressions: `${_row}`, `Row ${_row}: ${e.Name}`.
**Files:** `area.go` (`transformCell`)
**Tests:** _row in each loop, _col across columns, mixed expression with _row

## Enhancement 9: Differential Context Map Updates (Performance)
**Status:** Done
**Impact:** High — 17% speedup on 10K-row benchmarks
**What:** Replaced full map rebuild on every `runVar` change with differential in-place updates. Only the changed keys are updated in the cached map. Full rebuilds only happen when the data map itself changes (via `PutVar`/`RemoveVar`). For a 10K-row loop with 3 cells/row, this eliminates ~30K full map copies.
**Files:** `context.go` (differential caching via `mapDirtyKeys` + `mapNeedsFull` flags)
**Tests:** Differential update identity, nested loop save/restore, PutVar triggers rebuild

## Enhancement 10: Pre-allocated Slices/Maps (Performance)
**Status:** Done
**Impact:** Medium — reduced GC pressure for large templates
**What:** Pre-counted commented and formula cells during template loading. `GetCommentedCells()` and `GetFormulaCells()` now use pre-allocated slices. Also added unsigned int types to `toFloat64()` for correct numeric comparisons.
**Files:** `excelize_tx.go` (commentCellCount, formulaCellCount, pre-allocated result slices), `each.go` (uint types in toFloat64, pre-allocated filter slice)
**Tests:** Pre-allocation verification, toFloat64 all numeric types

## Enhancement 11: Structured Error Types
**Status:** Done
**Impact:** High — enables programmatic error handling at scale
**What:** Introduced `XLFillError` with `ErrorKind` (Template, Data, Runtime). All satisfy `error` and `errors.Unwrap()` for use with `errors.Is`/`errors.As`. Added `Warning` type and `WarningCollector` for non-fatal issues.
**Files:** `errors.go` (new — XLFillError, ErrorKind, Warning, WarningCollector, levenshtein, suggestCommand), `filler.go` (uses NewTemplateError), `xlfill.go` (uses NewRuntimeError)
**Tests:** Error type creation, errors.As/Is, ErrorKind.String(), Warning formatting

## Enhancement 12: Unknown Command Warnings (Strict Mode)
**Status:** Done
**Impact:** High — catches template typos that previously failed silently
**What:** Unknown `jx:` commands (e.g., `jx:eache`) now produce warnings with fuzzy-match suggestions (Levenshtein distance). `WithStrictMode(true)` turns these into errors. Default mode collects warnings via `Filler.Warnings()`.
**Files:** `errors.go` (levenshtein, suggestCommand), `filler.go` (unknown command handling, warnings), `command.go` (KnownNames), `options.go` (WithStrictMode)
**Tests:** Warning with "did you mean" suggestion, strict mode error, levenshtein distance, warnings reset between calls

## Enhancement 13: Context Variable Scope Validation
**Status:** Done
**Impact:** High — catches data/template mismatches before processing
**What:** `ValidateData(templatePath, data)` parses all expressions, extracts top-level variable names, cross-references against data map and command-provided variables (each var, repeat var), and reports missing variables as warnings.
**Files:** `validate.go` (ValidateData, collectCommandVars, checkDataRequirements, extractTopLevelVars), `xlfill.go` (top-level ValidateData function)
**Tests:** Missing data variable detection, command vars correctly excluded, keyword filtering

## Enhancement 14: Progress Reporting + Context Cancellation
**Status:** Done
**Impact:** Medium-High — critical for server-side and CLI use
**What:** `WithProgressFunc` reports rows processed during Fill. `WithContext(ctx)` enables cancellation/timeout. Every area/row processing step checks the context and calls the progress function.
**Files:** `options.go` (FillProgress, ProgressFunc, WithProgressFunc, WithContext), `area.go` (checkCancelled, reportProgress, ctx/progressFunc fields), `filler.go` (propagates ctx/debug/progressFunc to areas)
**Tests:** Progress callback invocation, context cancellation, context timeout

## Enhancement 15: Data Contract Validation
**Status:** Done
**Impact:** High — shifts errors left to development time
**What:** `ValidateData` checks expressions against provided data. Extracts top-level identifiers from expressions, filters out keywords and command-provided variables, reports mismatches. Works with nested commands (each, repeat, if, grid).
**Files:** `validate.go` (extractTopLevelVars, isExprKeyword, isIdentStart, isIdentPart, checkDataRequirements, collectCommandVars)
**Tests:** Missing variable detection, dot-access parsing, keyword filtering

## Enhancement 16: Debug/Trace Mode
**Status:** Done
**Impact:** High — dramatically reduces template debugging time
**What:** `WithDebugWriter(w)` produces structured trace output: area processing, command execution, each iterations, expression evaluations, and timing. `DebugTracer` also implements `AreaListener` for automatic cell counting.
**Files:** `debug.go` (new — DebugTracer with TraceArea, TraceCommand, TraceEachStart, TraceIteration, TraceExpr, TraceDone), `options.go` (WithDebugWriter), `filler.go` (debug propagation), `area.go` (debug integration), `each.go` (debug integration)
**Tests:** Trace output contains [area], [each], [iter], [done]; DebugTracer implements AreaListener; indent/dedent

## Enhancement 17: Compiled Template Reuse
**Status:** Done
**Impact:** High — major speedup for batch report generation
**What:** `Compile(templatePath)` returns a `CompiledTemplate` that caches the template bytes in memory. Each `Fill()` call creates a fresh transformer from the cached bytes — no filesystem I/O. Supports `Fill`, `FillBytes`, `FillWriter`. Validates the template during compilation.
**Files:** `compiled.go` (new — CompiledTemplate, Compile, Fill, FillBytes, FillWriter)
**Tests:** Multiple fills with different data, file output, invalid template, options propagation

## Enhancement 18: Style Callbacks (StyleListener)
**Status:** Done
**Impact:** Medium — eliminates need for raw excelize access for conditional styling
**What:** `StyleListener` interface with `StyleCell(target, value, ctx) *StyleOverride`. Supports Bold, Italic, FontColor, FillColor, FontSize. Propagated to all areas. Applied after cell transformation. Uses excelize `NewStyle` for actual styling.
**Files:** `listener.go` (StyleOverride, StyleListener), `area.go` (applyStyleOverrides, applyStyleToCell), `filler.go` (propagateStyleListeners)
**Tests:** Style listener invocation, StyleOverride field verification

## Enhancement 19: jx:repeat Command
**Status:** Done
**Impact:** Medium — eliminates dummy slice boilerplate
**What:** `jx:repeat(count="N" var="i" lastCell="C5")` repeats an area N times. Count can be a literal or expression. Optional `var` exposes 0-based index. Supports `direction="RIGHT"`. Zero count silently produces no output.
**Files:** `repeat.go` (new — RepeatCommand), `command.go` (registered in registry), `filler.go` (attachArea, getCommandArea, propagateListeners), `describe.go` (RepeatCommand attributes), `validate.go` (repeat command validation)
**Tests:** Basic repeat, expression count, direction RIGHT, zero count, missing count error, no area error, describe output, validation

## Enhancement 20: Transformer Interface Split
**Status:** Done
**Impact:** Medium — cleaner abstraction for custom implementations
**What:** Split `Transformer` into composable sub-interfaces: `CellReader`, `CellWriter`, `SheetManager`. The full `Transformer` embeds all three plus target tracking, image/merge/hyperlink, workbook properties, and I/O. Fully backward-compatible — `ExcelizeTransformer` satisfies all interfaces.
**Files:** `transformer.go` (CellReader, CellWriter, SheetManager interfaces)
**Tests:** Interface composition verification

## Bug Fixes (discovered during implementation)
1. **toFloat64 missing unsigned types** — `uint`, `uint8`, `uint16`, `uint32`, `uint64` were not handled, causing incorrect string-based comparisons for unsigned integers in `orderBy` and `compareValues`.
2. **parseOrderBy empty tokens** — `strings.Fields` could return empty result on malformed input; added guard.
3. **filterItems allocation** — Pre-allocates filtered slice with capacity estimate to reduce GC pressure.
4. **each.go error message** — Added type information to "items not iterable" error: `(got %T)`.
5. **RunVar double-close safety** — Verified Close() is safe to call multiple times (was already safe, added tests to document the contract).
