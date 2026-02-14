# xlfill Enhancement Plan

## Baseline
- 505 tests passing, 95.3% coverage
- All enhancements must maintain full backward compatibility (zero regression)

## Results
- **536 tests passing**, 94.0% coverage
- All 505 original tests pass (zero regression)
- 31 new tests added for enhancements
- 1 previously skipped test (multisheet) now passes

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
