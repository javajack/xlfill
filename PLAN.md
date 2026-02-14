# goxls Implementation Plan

## TDD-First, Phase-by-Phase Port of JXLS 3.0 to Go

**Principle**: Every feature is implemented test-first. No code is written without a failing test.
**Exclusion**: Database/JDBC connectivity (anti-pattern; violates separation of concerns).
**Quality gates**: Each phase must pass all its tests + all prior phase tests before proceeding.

---

## Phase 0: Project Bootstrap

### 0.1 Initialize Go Module
- Create `goxls/` directory with `go.mod` (`module github.com/user/goxls`)
- Add dependencies: `github.com/xuri/excelize/v2`, `github.com/expr-lang/expr`
- Add test dependency: `github.com/stretchr/testify` (assertions)
- Create package layout:
  ```
  goxls/
  ├── go.mod
  ├── goxls.go           # Public API
  ├── goxls_test.go      # Integration tests
  ├── cellref.go          # CellRef, AreaRef, Size
  ├── cellref_test.go
  ├── celldata.go         # CellData, CellType, FormulaStrategy
  ├── celldata_test.go
  ├── context.go          # Context, RunVar
  ├── context_test.go
  ├── expr.go             # Expression evaluator
  ├── expr_test.go
  ├── parse.go            # Comment/command parser
  ├── parse_test.go
  ├── command.go          # Command interface + registry
  ├── area.go             # Area interface + XlsArea
  ├── area_test.go
  ├── each.go             # EachCommand
  ├── each_test.go
  ├── condition.go        # IfCommand
  ├── condition_test.go
  ├── grid.go             # GridCommand
  ├── grid_test.go
  ├── image.go            # ImageCommand
  ├── image_test.go
  ├── mergecells.go       # MergeCellsCommand
  ├── mergecells_test.go
  ├── updatecell.go       # UpdateCellCommand
  ├── updatecell_test.go
  ├── transformer.go      # Transformer interface
  ├── excelize_tx.go      # excelize Transformer implementation
  ├── excelize_tx_test.go
  ├── formula.go          # FormulaProcessor interface + StandardFormulaProcessor
  ├── formula_test.go
  ├── filler.go           # Filler (template orchestrator)
  ├── filler_test.go
  ├── options.go          # Functional options
  └── testdata/           # Test .xlsx templates
  ```
- Single package `goxls` (flat, idiomatic Go for a focused library)
- Run `go mod tidy`, verify clean compile with `go build ./...`

### 0.2 CI/Quality Setup
- `Makefile` with targets: `test`, `test-race`, `bench`, `lint`, `cover`
- `go vet ./...` and `staticcheck` in lint target
- Test coverage threshold: 80% minimum, target 90%+
- Benchmark target for regression tracking

**Tests**: `go build ./...` passes, `go test ./...` passes (empty test files with package declaration)

**JXLS Parity**: N/A (infrastructure)

---

## Phase 1: Core Data Structures

### 1.1 CellRef — Cell Reference
**File**: `cellref.go`, `cellref_test.go`

**Implementation**:
```go
type CellRef struct {
    Sheet string // sheet name (empty = current sheet)
    Row   int    // 0-based row index
    Col   int    // 0-based column index
}
```

**Methods**:
- `NewCellRef(sheet string, row, col int) CellRef`
- `ParseCellRef(s string) (CellRef, error)` — parse "Sheet1!A1", "B5", "Sheet1!$A$1"
- `(c CellRef) String() string` — format as "Sheet1!A1"
- `(c CellRef) CellName() string` — format as "A1" (no sheet)
- `ColToName(col int) string` — 0→"A", 25→"Z", 26→"AA"
- `NameToCol(name string) (int, error)` — "A"→0, "AA"→26
- `(c CellRef) Equal(other CellRef) bool`

**Tests** (parity with JXLS `CellRef` + `CellDataJTest`):
- `TestParseCellRef_SimpleCell` — "A1" → {Row:0, Col:0}
- `TestParseCellRef_WithSheet` — "Sheet1!B5" → {Sheet:"Sheet1", Row:4, Col:1}
- `TestParseCellRef_AbsoluteRef` — "$A$1" → {Row:0, Col:0}
- `TestParseCellRef_MultiLetterCol` — "AA1" → {Row:0, Col:26}
- `TestParseCellRef_Invalid` — "!!!" → error
- `TestCellRef_String` — roundtrip: parse then format
- `TestColToName` — 0→"A", 25→"Z", 26→"AA", 702→"AAA"
- `TestNameToCol` — reverse of ColToName
- `TestCellRef_Equal` — same/different refs

### 1.2 AreaRef — Rectangular Area Reference
**File**: `cellref.go`, `cellref_test.go`

**Implementation**:
```go
type AreaRef struct {
    First CellRef
    Last  CellRef
}
```

**Methods**:
- `NewAreaRef(first, last CellRef) AreaRef`
- `ParseAreaRef(s string) (AreaRef, error)` — parse "A1:C5", "Sheet1!A1:C5"
- `(a AreaRef) String() string`
- `(a AreaRef) Size() Size`
- `(a AreaRef) Contains(ref CellRef) bool`
- `(a AreaRef) SheetName() string`

**Tests** (parity with JXLS `AreaRefContainsTest` — 9 boundary tests):
- `TestAreaRef_Contains_Inside` — cell inside area
- `TestAreaRef_Contains_TopLeft` — cell at top-left boundary
- `TestAreaRef_Contains_TopRight` — cell at top-right boundary
- `TestAreaRef_Contains_BottomLeft` — cell at bottom-left boundary
- `TestAreaRef_Contains_BottomRight` — cell at bottom-right boundary
- `TestAreaRef_Contains_Outside_Left` — cell outside left
- `TestAreaRef_Contains_Outside_Right` — cell outside right
- `TestAreaRef_Contains_Outside_Above` — cell above area
- `TestAreaRef_Contains_Outside_Below` — cell below area
- `TestAreaRef_Contains_DifferentSheet` — cell on different sheet
- `TestParseAreaRef` — "A1:C5" → correct first/last
- `TestAreaRef_Size` — correct width and height calculation

### 1.3 Size — Dimensions
**File**: `cellref.go`, `cellref_test.go`

**Implementation**:
```go
type Size struct {
    Width  int // columns
    Height int // rows
}
```

**Methods**:
- `(s Size) String() string`
- `(s Size) Add(other Size) Size`
- `(s Size) Minus(other Size) Size`

**Tests** (parity with JXLS `SizeTest`):
- `TestSize_String` — "(3x5)"
- `TestSize_Add` — (2,3) + (1,4) = (3,7)
- `TestSize_Minus` — (5,5) - (2,3) = (3,2)
- `TestSize_Zero` — zero size constant

### 1.4 CellData — Cell Information
**File**: `celldata.go`, `celldata_test.go`

**Implementation**:
```go
type CellType int
const (
    CellBlank CellType = iota
    CellString
    CellNumber
    CellBoolean
    CellDate
    CellFormula
    CellError
)

type FormulaStrategy int
const (
    FormulaDefault FormulaStrategy = iota
    FormulaByColumn
    FormulaByRow
)

type CellData struct {
    Ref              CellRef
    Value            interface{}
    Type             CellType
    Comment          string
    Formula          string
    EvalResult       interface{}
    TargetCellType   CellType
    FormulaStrategy  FormulaStrategy
    DefaultValue     string
    TargetPositions  []CellRef
    TargetParentArea []AreaRef
    EvalFormulas     []string
    StyleID          int
}
```

**Methods**:
- `(cd *CellData) AddTargetPos(ref CellRef)`
- `(cd *CellData) IsFormulaCell() bool`
- `(cd *CellData) Reset()`

**Tests** (parity with JXLS `CellDataJTest`):
- `TestCellData_Construction` — create with all fields
- `TestCellData_TargetPositions` — add/track multiple targets
- `TestCellData_IsFormulaCell` — formula vs non-formula
- `TestCellData_Reset` — clears target positions

**Phase 1 Quality Gate**: `go test ./... -count=1` passes, all ~25 tests green.

---

## Phase 2: Expression Evaluator

### 2.1 Expression Evaluator
**File**: `expr.go`, `expr_test.go`

**Implementation**:
```go
type ExpressionEvaluator interface {
    Evaluate(expression string, data map[string]any) (any, error)
    IsConditionTrue(condition string, data map[string]any) (bool, error)
}

type exprEvaluator struct{}  // backed by expr-lang/expr
```

**Features**:
- Property access on structs (exported fields) and maps
- Arithmetic, comparison, logical, ternary operators
- Built-in functions: `len()` for size/empty checks
- Nil-safe evaluation (return nil, not panic)
- Compiled expression cache (sync.Map for concurrency)

**Tests** (parity with JXLS `JexlExpressionEvaluatorTest`):
- `TestExpr_SimpleProperty` — `e.Name` on struct
- `TestExpr_NestedProperty` — `e.Address.City`
- `TestExpr_MapAccess` — `data["key"]` on map
- `TestExpr_Arithmetic` — `e.A + e.B`
- `TestExpr_Comparison` — `e.Payment > 2000`
- `TestExpr_LogicalAnd` — `e.A > 0 && e.B < 100`
- `TestExpr_LogicalOr` — `e.A > 0 || e.B < 100`
- `TestExpr_Ternary` — `e.Flag ? "Yes" : "No"`
- `TestExpr_LenFunction` — `len(list) == 0`, `len(list)`
- `TestExpr_NullVariable` — unknown variable returns nil, no panic
- `TestExpr_StringConcat` — `"Hello " + e.Name`
- `TestExpr_ErrorExpression` — invalid syntax returns error
- `TestExpr_IsConditionTrue` — boolean expression evaluation
- `TestExpr_IsConditionFalse` — false condition
- `TestExpr_NilCondition` — nil treated as false (JXLS v3 behavior)
- `TestExpr_SliceAccess` — `items[0]`
- `TestExpr_StructAndMap_Mixed` — struct with map field
- `TestExpr_ConcurrencySafe` — parallel evaluation (goroutines)

### 2.2 Expression Notation Parser
**File**: `expr.go`, `expr_test.go`

**Implementation**:
- Parse `${...}` patterns from cell value strings
- Support configurable notation (e.g., `{{...}}`)
- Handle multiple expressions in one string: `"Name: ${e.Name}, Age: ${e.Age}"`
- Handle expression-only cells: `"${e.Name}"` → evaluate to typed value
- Handle mixed content: `"Total: ${e.Total}"` → always string

**Methods**:
- `ParseExpressions(value string, begin, end string) []ExpressionSegment`
- `EvaluateCellValue(value string, ctx *Context) (any, CellType, error)`

**Tests**:
- `TestParseExpressions_Single` — `"${e.Name}"` → 1 expression
- `TestParseExpressions_Multiple` — `"${e.First} ${e.Last}"` → 2 expressions + literals
- `TestParseExpressions_NoExpr` — `"Hello"` → 1 literal
- `TestParseExpressions_CustomNotation` — `"{{e.Name}}"` with `{{`, `}}`
- `TestParseExpressions_Nested` — `"${e.Map["key"]}"` — braces inside
- `TestEvaluateCellValue_TypedResult` — `"${e.Payment}"` → number type
- `TestEvaluateCellValue_StringResult` — `"Name: ${e.Name}"` → string type
- `TestEvaluateCellValue_BoolResult` — `"${e.Active}"` → bool type
- `TestEvaluateCellValue_NilResult` — `"${e.Missing}"` → blank type

**Phase 2 Quality Gate**: All Phase 1 + Phase 2 tests pass. Expression evaluation handles all JXLS expression patterns.

---

## Phase 3: Context & Variable Management

### 3.1 Context
**File**: `context.go`, `context_test.go`

**Implementation**:
```go
type Context struct {
    data             map[string]any
    runVars          map[string]any
    evaluator        ExpressionEvaluator
    notationBegin    string
    notationEnd      string
    updateCellData   bool
    clearCells       bool
}
```

**Methods**:
- `NewContext(data map[string]any, opts ...ContextOption) *Context`
- `(c *Context) GetVar(name string) any`
- `(c *Context) PutVar(name string, value any)`
- `(c *Context) RemoveVar(name string)`
- `(c *Context) ContainsVar(name string) bool`
- `(c *Context) ToMap() map[string]any` — merge data + runVars
- `(c *Context) Evaluate(expression string) (any, error)`
- `(c *Context) IsConditionTrue(condition string) (bool, error)`
- `(c *Context) EvaluateCellValue(cellValue string) (any, CellType, error)`

### 3.2 RunVar — Scoped Loop Variables
**File**: `context.go`, `context_test.go`

**Implementation**:
```go
type RunVar struct {
    ctx      *Context
    varName  string
    oldValue any
    hadOld   bool
    idxName  string
    oldIdx   any
    hadIdx   bool
}
```

**Methods**:
- `NewRunVar(ctx *Context, varName string) *RunVar`
- `NewRunVarWithIndex(ctx *Context, varName, idxName string) *RunVar`
- `(rv *RunVar) Set(value any)`
- `(rv *RunVar) SetWithIndex(value any, index int)`
- `(rv *RunVar) Close()` — restore old values (defer-friendly)

**Tests** (parity with JXLS `ContextTest` + `RunVar` behavior):
- `TestContext_PutGetVar` — put and get variable
- `TestContext_RemoveVar` — remove variable
- `TestContext_ContainsVar` — check existence
- `TestContext_ToMap` — merge data + runVars
- `TestContext_Evaluate` — evaluate expression through context
- `TestContext_IsConditionTrue` — boolean condition
- `TestContext_RunVarScope` — set var, close restores old value
- `TestContext_RunVarScope_NewVar` — set var that didn't exist, close removes it
- `TestContext_RunVarWithIndex` — set var + index, close restores both
- `TestContext_RunVarNested` — nested runvars (outer loop var preserved)
- `TestContext_EvaluateCellValue_Expression` — `"${e.Name}"` via context
- `TestContext_EvaluateCellValue_Mixed` — `"Total: ${sum}"` via context

**Phase 3 Quality Gate**: All Phase 1-3 tests pass. Context correctly manages variable scoping for nested loops.

---

## Phase 4: Transformer (excelize Integration)

### 4.1 Transformer Interface
**File**: `transformer.go`

**Implementation**:
```go
type Transformer interface {
    // Cell operations
    GetCellData(ref CellRef) (*CellData, error)
    GetCommentedCells() []*CellData
    Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error
    ClearCell(ref CellRef) error
    SetFormula(ref CellRef, formula string) error

    // Row/column operations
    GetSheetData(sheet string) *SheetData
    GetColumnWidth(sheet string, col int) float64
    GetRowHeight(sheet string, row int) float64
    SetColumnWidth(sheet string, col int, width float64) error
    SetRowHeight(sheet string, row int, height float64) error

    // Tracking
    GetFormulaCells() []*CellData
    GetTargetCellRef(src CellRef) []CellRef
    ResetTargetCellRefs()

    // Sheet operations
    DeleteSheet(name string) error
    SetHidden(name string, hidden bool) error
    GetSheetNames() []string

    // Image/merge (POI-equivalent)
    AddImage(sheet string, cell string, imgBytes []byte, imgType string, scaleX, scaleY float64) error
    MergeCells(sheet, topLeft, bottomRight string) error

    // I/O
    Write(w io.Writer) error
}
```

### 4.2 SheetData, RowData
**File**: `transformer.go`

```go
type SheetData struct {
    Name         string
    ColumnWidths map[int]float64
    Rows         map[int]*RowData
}

type RowData struct {
    Height float64
    Cells  map[int]*CellData
}
```

### 4.3 ExcelizeTransformer — excelize Implementation
**File**: `excelize_tx.go`, `excelize_tx_test.go`

**Implementation**:
- Opens template with `excelize.OpenFile()`
- Reads ALL cell data, comments, styles, formulas into memory (like PoiTransformer)
- Caches style IDs for preservation during cell writes
- Tracks source→target cell mappings for formula processing
- Handles row height / column width copying

**Key internal methods**:
- `readAllCellData()` — populate SheetData from template
- `transformCell(src, target CellRef, ctx *Context)` — evaluate + write cell
- `copyCellStyle(src, target CellRef)` — cache and restore style
- `writeValue(sheet, cell string, value any, cellType CellType)` — typed write

**Tests** (parity with JXLS `PoiTransformer` behavior):
- `TestTransformer_OpenTemplate` — open existing .xlsx, read cells
- `TestTransformer_GetCellData` — read cell value, type, comment
- `TestTransformer_GetCommentedCells` — find all cells with comments
- `TestTransformer_Transform_StringValue` — copy string cell with style
- `TestTransformer_Transform_NumberValue` — copy number cell
- `TestTransformer_Transform_DateValue` — copy date cell
- `TestTransformer_Transform_FormulaCell` — copy formula cell
- `TestTransformer_Transform_BlankCell` — copy blank cell
- `TestTransformer_Transform_PreservesStyle` — style preserved after value change
- `TestTransformer_ClearCell` — clear cell content
- `TestTransformer_SetFormula` — set formula on cell
- `TestTransformer_TrackTargetCellRef` — source→target mapping
- `TestTransformer_Write` — write to output stream
- `TestTransformer_DeleteSheet` — delete sheet
- `TestTransformer_SetHidden` — hide sheet
- `TestTransformer_ColumnWidth` — read/copy column width
- `TestTransformer_RowHeight` — read/copy row height
- `TestTransformer_MergeCells` — merge cell range
- `TestTransformer_AddImage` — add image to sheet

**Test templates needed** (in `testdata/`):
- `testdata/transform_basic.xlsx` — basic cells with various types + comments
- `testdata/transform_styled.xlsx` — cells with formatting (bold, colors, borders, number formats)
- `testdata/transform_formula.xlsx` — cells with formulas
- `testdata/transform_merged.xlsx` — cells with merged regions
- `testdata/transform_image.xlsx` — template for image insertion

**Phase 4 Quality Gate**: All Phase 1-4 tests pass. Transformer correctly reads templates and writes cells preserving styles.

---

## Phase 5: Template Parser (Comment → Commands)

### 5.1 Command Interface & Registry
**File**: `command.go`

**Implementation**:
```go
type Command interface {
    Name() string
    Areas() []*Area
    AddArea(area *Area)
    ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error)
    Reset()
    ShiftMode() string
    SetShiftMode(mode string)
    LockRange() bool
    SetLockRange(lock bool)
}

// BaseCommand provides common implementation
type BaseCommand struct {
    areas     []*Area
    shiftMode string
    lockRange bool
}

// CommandRegistry maps command names to constructors
type CommandRegistry struct {
    commands map[string]func() Command
}
```

### 5.2 Comment Parser
**File**: `parse.go`, `parse_test.go`

**Implementation**:
- Parse `jx:COMMAND(attr1="val1" attr2="val2" lastCell="REF")` from cell comments
- Support multiple commands per comment (one per line)
- Support `areas=["A1:C1","A2:C2"]` attribute
- Handle Unicode quote characters (LibreOffice compatibility)
- Handle multi-line values (SQL feature parity, though we exclude JDBC the parser must handle multi-line attribute values)
- `jx:params` handled specially (no command class, attributes applied to CellData)

**Types**:
```go
type ParsedCommand struct {
    Name     string
    Attrs    map[string]string  // ordered attributes
    LastCell CellRef
    Areas    []AreaRef          // optional areas attribute
    CellRef  CellRef            // cell containing the comment
}
```

**Methods**:
- `ParseComment(comment string, cellRef CellRef) ([]ParsedCommand, error)`
- `ParseAttribute(attrStr string) map[string]string`
- `IsCommand(line string) bool` — starts with "jx:" and not "jx:params"
- `IsParams(line string) bool` — starts with "jx:params"
- `ParseParams(line string) (formulaStrategy FormulaStrategy, defaultValue string, err error)`

**Tests** (parity with JXLS `XlsCommentAreaBuilder` + `LiteralsExtractorTest`):
- `TestParseComment_SimpleEach` — `jx:each(items="list" var="e" lastCell="C2")`
- `TestParseComment_IfWithAreas` — `jx:if(condition="x>1" lastCell="C2" areas=["A2:C2","A3:C3"])`
- `TestParseComment_Area` — `jx:area(lastCell="D10")`
- `TestParseComment_Grid` — `jx:grid(headers="h" data="d" areas=["A1:A1","A2:A2"] lastCell="A2")`
- `TestParseComment_Image` — `jx:image(src="img" imageType="PNG" lastCell="A2")`
- `TestParseComment_MergeCells` — `jx:mergeCells(lastCell="D2" cols="4" rows="2")`
- `TestParseComment_UpdateCell` — `jx:updateCell(lastCell="E4" updater="myUpdater")`
- `TestParseComment_MultipleCommands` — two commands in one comment (two lines)
- `TestParseComment_EachAllAttrs` — all each attributes: items, var, varIndex, direction, select, groupBy, groupOrder, orderBy, multisheet
- `TestParseComment_WithCommas` — `jx:each(items="list", var="e", lastCell="C2")` (optional commas)
- `TestParseComment_UnicodeQuotes` — LibreOffice smart quotes
- `TestParseComment_WhitespaceVariants` — extra spaces around = and in values
- `TestParseComment_SheetInLastCell` — `lastCell="Sheet2!A5"`
- `TestParseComment_InvalidCommand` — missing lastCell → error
- `TestParseComment_UnknownCommand` — `jx:unknown(...)` → error or skip
- `TestParseComment_EmptyComment` — "" → empty result
- `TestParseComment_NonJxComment` — "This is a note" → empty result
- `TestIsParams` — `jx:params(defaultValue="1")` → true
- `TestParseParams_DefaultValue` — extract defaultValue
- `TestParseParams_FormulaStrategy` — extract formulaStrategy="BY_COLUMN"
- `TestParseParams_Both` — both attributes together
- `TestParseComment_MultiLine` — multi-line attribute values

### 5.3 Area Builder
**File**: `parse.go`, `parse_test.go`

**Implementation**:
- Scans all commented cells from Transformer
- Finds `jx:area` commands → creates root Areas
- Finds other commands → nests them within their containing Area
- Applies `jx:params` to CellData directly
- Returns list of root Areas ready for processing

**Methods**:
- `BuildAreas(transformer Transformer, clearTemplateCells bool) ([]*Area, error)`

**Tests**:
- `TestBuildAreas_SingleArea` — one jx:area with one jx:each
- `TestBuildAreas_NestedCommands` — jx:each containing jx:if
- `TestBuildAreas_MultipleAreas` — multiple jx:area on different sheets
- `TestBuildAreas_ParamsApplied` — jx:params sets CellData attributes
- `TestBuildAreas_NoAreaCommand` — template without jx:area → error
- `TestBuildAreas_CustomCommands` — user-registered commands found

**Phase 5 Quality Gate**: All Phase 1-5 tests pass. Templates can be parsed into Area/Command hierarchy.

---

## Phase 6: Area Processing Engine

### 6.1 Area & XlsArea
**File**: `area.go`, `area_test.go`

**Implementation**:
```go
type Area struct {
    startCell   CellRef
    size        Size
    commands    []*CommandData
    transformer Transformer
    parent      Command
    cellRange   *CellRange  // 2D tracking array
}

type CommandData struct {
    SourceStart CellRef
    SourceSize  Size
    Start       CellRef  // current (shifted) position
    Size        Size
    Command     Command
}
```

**Core algorithm** (port of XlsArea.applyAt):
1. Create CellRange (2D state array: excluded, transformed, cleared)
2. Exclude locked command ranges
3. Transform static cells above first command
4. For each command:
   a. Execute command.ApplyAt() → get resulting Size
   b. Calculate height/width change
   c. If height changed: shift dependent cells/commands vertically
   d. If width changed: shift dependent cells/commands horizontally
   e. Transform static cells between commands
5. Transform remaining static cells
6. Track formula cells for later processing
7. Return final area size

**CellRange internal**:
```go
type CellRange struct {
    width, height int
    excluded      [][]bool
    transformed   [][]bool
}
```

**Methods**:
- `(a *Area) ApplyAt(cellRef CellRef, ctx *Context) (Size, error)`
- `(a *Area) ProcessFormulas(fp FormulaProcessor)`
- `(a *Area) ClearCells()`
- `(a *Area) transformCell(srcRef, targetRef CellRef, ctx *Context) error`
- `(a *Area) shiftVertical(startRow, heightChange int, affectedCols ...int)`
- `(a *Area) shiftHorizontal(startCol, widthChange int, affectedRows ...int)`

**Tests** (parity with JXLS `XlsAreaTest`):
- `TestArea_ApplyAt_StaticCells` — area with no commands, just expressions
- `TestArea_ApplyAt_SingleCommand` — area with one jx:each
- `TestArea_ApplyAt_MultipleCommands` — area with two sequential commands
- `TestArea_ApplyAt_CommandExpansion` — command adds rows, subsequent cells shift
- `TestArea_ApplyAt_CommandContraction` — command removes rows (empty list)
- `TestArea_ApplyAt_MultiSheet` — area spanning multiple sheets
- `TestArea_ShiftDown` — cells below command shift down on expansion
- `TestArea_ShiftRight` — cells right of command shift right
- `TestArea_ClearCells` — template cells cleared after processing
- `TestArea_ExcludedCells` — locked range cells not re-transformed
- `TestArea_NestedCommands` — commands within commands (each > if)
- `TestArea_CascadeShift` — expansion of one command shifts all below

**Phase 6 Quality Gate**: All Phase 1-6 tests pass. Area processing correctly transforms cells and handles shifts.

---

## Phase 7: EachCommand (Basic)

### 7.1 EachCommand — Basic Iteration (DOWN direction)
**File**: `each.go`, `each_test.go`

**Implementation**:
```go
type EachCommand struct {
    BaseCommand
    Items     string  // expression for collection
    Var       string  // loop variable name
    Direction string  // "DOWN" (default) or "RIGHT"
    VarIndex  string  // optional index variable name
    // Advanced (Phase 10): Select, GroupBy, GroupOrder, OrderBy, MultiSheet
}
```

**Algorithm** (basic):
1. Evaluate `Items` expression → get slice/array
2. For each item:
   a. Create RunVar with Var (and VarIndex if set)
   b. Defer RunVar.Close()
   c. Calculate target cell: startRow + accumulatedHeight (DOWN)
   d. Call area.ApplyAt(targetCell, ctx) → size
   e. Accumulate height
3. Return total Size

**Tests** (parity with JXLS `EachTest` basic):
- `TestEachCommand_BasicList` — 3 employees, verify all rows populated
- `TestEachCommand_EmptyList` — empty slice → zero size, no rows
- `TestEachCommand_NilList` — nil items → zero size, no rows
- `TestEachCommand_SingleItem` — 1 item → 1 row
- `TestEachCommand_LargeList` — 100 items → 100 rows
- `TestEachCommand_VarIndex` — varIndex accessible in expressions
- `TestEachCommand_PreservesFormatting` — styles copied to each row
- `TestEachCommand_MultiColumnTemplate` — template with 3+ columns per row
- `TestEachCommand_ExpressionInCells` — ${e.Name}, ${e.Age} evaluated correctly
- `TestEachCommand_NumberTypes` — int, float64, correctly typed in output
- `TestEachCommand_DateTypes` — time.Time correctly written
- `TestEachCommand_NilFieldValue` — nil field → blank cell
- `TestEachCommand_NestedStruct` — ${e.Address.City}

**Test template**: `testdata/each_basic.xlsx`
```
A1: "Name"           B1: "Age"        C1: "Salary"     [comment: jx:area(lastCell="C2")]
A2: "${e.Name}"      B2: "${e.Age}"   C2: "${e.Salary}" [comment: jx:each(items="employees" var="e" lastCell="C2")]
```

**Phase 7 Quality Gate**: Basic each works end-to-end with template file.

---

## Phase 8: IfCommand

### 8.1 IfCommand — Conditional Rendering
**File**: `condition.go`, `condition_test.go`

**Implementation**:
```go
type IfCommand struct {
    BaseCommand
    Condition string  // boolean expression
    IfArea    *Area   // rendered when true
    ElseArea  *Area   // rendered when false (optional)
}
```

**Algorithm**:
1. Evaluate Condition expression → bool
2. If true: apply IfArea at target cell
3. If false and ElseArea exists: apply ElseArea at target cell
4. If false and no ElseArea: return zero size
5. Return resulting size

**Tests** (parity with JXLS `IfTest`):
- `TestIfCommand_True` — condition true → if area rendered
- `TestIfCommand_False` — condition false → else area rendered
- `TestIfCommand_FalseNoElse` — condition false, no else → nothing rendered
- `TestIfCommand_NilCondition` — nil treated as false (JXLS v3 behavior)
- `TestIfCommand_InsideEach` — if inside each loop (different conditions per row)
- `TestIfCommand_PreservesFormatting` — styles preserved in conditional areas

**Test template**: `testdata/if_basic.xlsx`
```
A1: [jx:area] ... 
A2: ${e.Name}  B2: ${e.Payment}  C2: "High" [jx:if(condition="e.Payment > 2000" lastCell="C2")]
A3: ${e.Name}  B3: ${e.Payment}  C3: "Low"  [else area referenced in jx:if areas attr]
```

**Phase 8 Quality Gate**: If/else works standalone and nested inside each.

---

## Phase 9: End-to-End Integration (Filler)

### 9.1 Filler — Template Processing Orchestrator
**File**: `filler.go`, `filler_test.go`

**Implementation**:
```go
type Filler struct {
    templatePath     string
    templateReader   io.Reader
    options          *Options
}
```

**Algorithm** (port of JxlsTemplateFiller.fill):
1. Open template → create Transformer
2. Configure Transformer (column/row props, sheet creator)
3. Build Areas (parse comments → Area/Command hierarchy)
4. Create Context with user data
5. For each Area: area.ApplyAt(startCell, ctx)
6. If FormulaProcessor configured: processFormulas for each Area
7. Pre-write actions (delete/hide template sheets, recalculation flags)
8. Write output

### 9.2 Public API
**File**: `goxls.go`

```go
// Simple API
func Fill(templatePath, outputPath string, data map[string]any) error
func FillBytes(templatePath string, data map[string]any) ([]byte, error)
func FillReader(template io.Reader, output io.Writer, data map[string]any) error

// Builder API
func NewFiller(opts ...Option) *Filler
func (f *Filler) Fill(data map[string]any, outputPath string) error
func (f *Filler) FillWriter(data map[string]any, w io.Writer) error
func (f *Filler) FillBytes(data map[string]any) ([]byte, error)
```

### 9.3 Functional Options
**File**: `options.go`

```go
type Option func(*Options)

func WithTemplate(path string) Option
func WithTemplateReader(r io.Reader) Option
func WithExpressionNotation(begin, end string) Option
func WithFormulaProcessor(fp FormulaProcessor) Option
func WithCommand(name string, factory func() Command) Option
func WithClearTemplateCells(clear bool) Option
func WithKeepTemplateSheet(keep bool) Option
func WithHideTemplateSheet(hide bool) Option
func WithPreWrite(fn func(Transformer) error) Option
func WithIgnoreColumnProps(ignore bool) Option
func WithIgnoreRowProps(ignore bool) Option
```

**Tests** — Full integration tests:
- `TestFill_BasicEach` — employees template → populated output
- `TestFill_EachWithIf` — each + if combined
- `TestFill_EmptyList` — empty data produces header-only output
- `TestFill_MultipleAreas` — multiple jx:area on one sheet
- `TestFill_OutputToFile` — write to file path
- `TestFill_OutputToWriter` — write to io.Writer
- `TestFill_OutputToBytes` — get byte array
- `TestFill_PreservesFormatting` — bold, colors, number formats preserved
- `TestFill_ClearsTemplateCells` — template expressions removed from output
- `TestFill_CustomNotation` — `{{...}}` notation
- `TestFill_TemplateFromReader` — template as io.Reader
- `TestFill_InvalidTemplate` — non-existent file → error
- `TestFill_NilData` — nil data map → error
- `TestFill_MapData` — map[string]any as item data (no struct)
- `TestFill_StructData` — struct as item data
- `TestFill_MixedData` — some map, some struct

**Test template**: `testdata/integration_basic.xlsx`

**Phase 9 Quality Gate**: Full end-to-end pipeline works. Users can call `goxls.Fill()` and get populated Excel output.

---

## Phase 10: Advanced EachCommand Features

### 10.1 Direction RIGHT
**File**: `each.go`, `each_test.go`

**Implementation**: When direction="RIGHT", iterate horizontally instead of vertically.
- Target cell: startCol + accumulatedWidth (instead of startRow + accumulatedHeight)
- Accumulate widths, track max height

**Tests** (parity with JXLS `DirectionRightTest`):
- `TestEachCommand_DirectionRight` — horizontal expansion
- `TestEachCommand_DirectionRight_TwoColumns` — 2-column template area
- `TestEachCommand_DirectionRight_FourColumns` — 4-column template area
- `TestEachCommand_DirectionRight_WithFormulas` — column sums

**Test template**: `testdata/each_direction_right.xlsx`

### 10.2 Select (Filtering)
**File**: `each.go`, `each_test.go`

**Implementation**: Evaluate `Select` expression per item, skip items where false.

**Tests** (parity with JXLS `SelectTest`):
- `TestEachCommand_Select` — filter items by expression
- `TestEachCommand_Select_AllFiltered` — all items filtered → empty
- `TestEachCommand_Select_NoneFiltered` — no items filtered → all present
- `TestEachCommand_Select_ComplexExpression` — `e.Payment > 2000 && e.Active`

**Test template**: `testdata/each_select.xlsx`

### 10.3 OrderBy (Sorting)
**File**: `each.go`, `each_test.go`

**Implementation**:
- Parse orderBy string: `"e.Name ASC, e.Payment DESC"`
- Remove var prefix, extract property name + direction
- Sort collection using reflection-based comparator
- Support ASC, DESC, ASC_ignoreCase, DESC_ignoreCase

**Tests** (parity with JXLS `EachTest.orderBy` + `OrderByComparatorTest`):
- `TestEachCommand_OrderBy_Asc` — ascending sort
- `TestEachCommand_OrderBy_Desc` — descending sort
- `TestEachCommand_OrderBy_MultiProperty` — sort by name then payment
- `TestEachCommand_OrderBy_IgnoreCase` — case-insensitive sort
- `TestEachCommand_OrderBy_NilValues` — nil values handled (sorted last)
- `TestOrderByComparator_SingleAsc` — comparator unit test
- `TestOrderByComparator_SingleDesc` — descending comparator
- `TestOrderByComparator_MultiField` — multi-field comparator

**Test template**: `testdata/each_orderby.xlsx`

### 10.4 GroupBy/GroupOrder (Grouping)
**File**: `each.go`, `each_test.go`

**Implementation**:
- Group items by property value into `GroupData{Item, Items}`
- GroupOrder: ASC, DESC, ASC_ignoreCase, DESC_ignoreCase
- Select applied BEFORE grouping (new behavior, default)
- Access: `${g.Item.Department}`, iterate `${g.Items}`

```go
type GroupData struct {
    Item  any   // the group key value
    Items []any // items in this group
}
```

**Tests** (parity with JXLS `GroupByTest` — 5 variants):
- `TestEachCommand_GroupBy_Asc` — groups sorted ascending
- `TestEachCommand_GroupBy_Desc` — groups sorted descending
- `TestEachCommand_GroupBy_AscIgnoreCase` — case-insensitive ascending
- `TestEachCommand_GroupBy_DescIgnoreCase` — case-insensitive descending
- `TestEachCommand_GroupBy_WithSelect` — select + groupBy combined
- `TestEachCommand_GroupBy_NilGroupKey` — nil group key handled
- `TestEachCommand_GroupBy_NestedEach` — each inside grouped each

**Test templates**: `testdata/each_groupby_asc.xlsx`, `testdata/each_groupby_desc.xlsx`, etc.

### 10.5 MultiSheet
**File**: `each.go`, `each_test.go`

**Implementation**:
- `MultiSheet` attribute: variable name containing []string of sheet names
- Each iteration creates/uses a different sheet
- Overrides direction-based movement
- Template sheet can be deleted/hidden after processing
- Safe sheet name generation (Excel restrictions: 31 chars, no []:*?/\ chars)

**Tests** (parity with JXLS `MultiSheetTest`):
- `TestEachCommand_MultiSheet` — create multiple sheets from template
- `TestEachCommand_MultiSheet_DeleteTemplate` — template sheet deleted
- `TestEachCommand_MultiSheet_HideTemplate` — template sheet hidden
- `TestEachCommand_MultiSheet_DuplicateNames` — duplicate names made unique
- `TestEachCommand_MultiSheet_SpecialChars` — invalid chars replaced
- `TestSafeSheetName` — sheet name sanitization unit tests

**Test template**: `testdata/each_multisheet.xlsx`

### 10.6 Scalar/Primitive Arrays
**File**: `each.go`, `each_test.go`

**Tests** (parity with JXLS `ScalarsTest`):
- `TestEachCommand_IntSlice` — iterate []int
- `TestEachCommand_Float64Slice` — iterate []float64
- `TestEachCommand_StringSlice` — iterate []string

**Phase 10 Quality Gate**: All advanced each features work. All Phase 1-10 tests pass.

---

## Phase 11: Formula Processing

### 11.1 FormulaProcessor Interface
**File**: `formula.go`

```go
type FormulaProcessor interface {
    ProcessAreaFormulas(transformer Transformer, area *Area)
}
```

### 11.2 StandardFormulaProcessor
**File**: `formula.go`, `formula_test.go`

**Implementation** (port of StandardFormulaProcessor):
1. For each formula cell in transformer:
   - If cell belongs to this area:
     - For each target position of this formula cell:
       - Get source formula
       - For each cell reference in formula:
         - Look up target positions of referenced cell
         - Filter by target area (internal vs external refs)
         - Apply BY_COLUMN/BY_ROW strategy if set
         - Replace source ref with target ref(s) in formula
       - Handle jointed cell references (`U_(...)`)
       - Handle SUM with 255+ args (convert to addition)
       - Apply default value for empty refs
       - Write updated formula to target cell

**Formula cell reference parsing**:
- Regex to find cell references in formula strings (e.g., "SUM(C2:C2)")
- Handle sheet-qualified refs: "Sheet1!A1"
- Handle absolute refs: "$A$1"
- Handle range refs: "A1:A10"

**Tests** (parity with JXLS `CreateTargetCellRefTest` + `FormulaProcessorsTest` + `FastFormulaProcessorTest`):
- `TestFormulaProcessor_SimpleSum` — SUM(C2:C2) → SUM(C2:C5) after 4-row expansion
- `TestFormulaProcessor_SingleCellRef` — A2 → A2 (no expansion)
- `TestFormulaProcessor_HorizontalRange` — A2:D2 → A2:G2 (RIGHT expansion)
- `TestFormulaProcessor_VerticalRange` — A2:A2 → A2:A10
- `TestFormulaProcessor_RectangleRange` — 2D expansion
- `TestFormulaProcessor_GapInRange` — A2,A4 (A3 excluded by if)
- `TestFormulaProcessor_ExternalRef` — ref outside area (preserved)
- `TestFormulaProcessor_ExternalRefReplicated` — external ref in outer loop
- `TestFormulaProcessor_ByColumnStrategy` — only match same column
- `TestFormulaProcessor_ByRowStrategy` — only match same row
- `TestFormulaProcessor_DefaultValue` — removed ref → default value
- `TestFormulaProcessor_CustomDefaultValue` — jx:params(defaultValue="1")
- `TestFormulaProcessor_SumOver255` — SUM with 256+ args → addition chain
- `TestFormulaProcessor_CrossSheetRef` — Sheet2!A1 reference
- `TestFormulaProcessor_NoProcessor` — formula processing disabled
- `TestFormulaProcessor_NestedLoopFormulas` — formulas in nested each

**Test templates**: `testdata/formula_basic.xlsx`, `testdata/formula_nested.xlsx`

### 11.3 FastFormulaProcessor (Optional Optimization)
**File**: `formula.go`

Simplified processor: only handles simple SUM-style formula expansion within single areas.
No external ref support, no jointed refs, no cross-sheet.

**Tests**:
- `TestFastFormulaProcessor_SimpleCase` — basic SUM expansion
- `TestFastFormulaProcessor_Limitation` — complex case handled gracefully

**Phase 11 Quality Gate**: Formulas auto-update after template expansion. All Phase 1-11 tests pass.

---

## Phase 12: GridCommand

### 12.1 GridCommand — Dynamic Grid
**File**: `grid.go`, `grid_test.go`

**Implementation**:
```go
type GridCommand struct {
    BaseCommand
    Headers     string  // expression for header values
    Data        string  // expression for data rows
    Props       string  // comma-separated property names (for object data)
    FormatCells string  // type-to-cell format mapping
    HeaderArea  *Area
    BodyArea    *Area
}
```

**Algorithm**:
1. Evaluate Headers → []any
2. Evaluate Data → []any (rows)
3. For each header: render HeaderArea horizontally
4. For each data row:
   - If row is slice/array: render each element with BodyArea
   - If row is struct/map with Props: extract properties, render each

**Tests** (parity with JXLS `GridTest`):
- `TestGridCommand_BasicGrid` — headers + data rendered
- `TestGridCommand_NilHeaders` — nil headers → 0 columns
- `TestGridCommand_NilData` — nil data → 0 rows
- `TestGridCommand_ObjectDataWithProps` — struct data with property names
- `TestGridCommand_ArrayData` — [][]any data
- `TestGridCommand_FormatCells` — type-based formatting

**Test template**: `testdata/grid_basic.xlsx`

**Phase 12 Quality Gate**: Grid command works. All Phase 1-12 tests pass.

---

## Phase 13: ImageCommand

### 13.1 ImageCommand — Image Embedding
**File**: `image.go`, `image_test.go`

**Implementation**:
```go
type ImageCommand struct {
    BaseCommand
    Src       string  // expression returning []byte
    ImageType string  // PNG, JPEG, etc. (default: PNG)
    ScaleX    float64 // width scale (default: 1.0)
    ScaleY    float64 // height scale (default: 1.0)
}
```

**Tests** (parity with JXLS `ImageTest`):
- `TestImageCommand_PNG` — insert PNG image
- `TestImageCommand_JPEG` — insert JPEG image
- `TestImageCommand_WithScaling` — scaled image
- `TestImageCommand_NilBytes` — nil image data → skip gracefully

**Test template**: `testdata/image_basic.xlsx`
**Test asset**: `testdata/test_image.png`

**Phase 13 Quality Gate**: Images can be inserted. All Phase 1-13 tests pass.

---

## Phase 14: MergeCellsCommand

### 14.1 MergeCellsCommand
**File**: `mergecells.go`, `mergecells_test.go`

**Implementation**:
```go
type MergeCellsCommand struct {
    BaseCommand
    Cols    string  // number of columns to merge (expression)
    Rows    string  // number of rows to merge (expression)
    MinCols string  // minimum cols for merge
    MinRows string  // minimum rows for merge
}
```

**Tests** (parity with JXLS `AreaColumnMergeTest` + `MergeCellsCommand`):
- `TestMergeCellsCommand_Basic` — merge specified range
- `TestMergeCellsCommand_Dynamic` — cols/rows from expressions
- `TestMergeCellsCommand_MinThreshold` — merge only if above minimum
- `TestMergeCellsCommand_InsideEach` — merge within loop iterations

**Test template**: `testdata/mergecells_basic.xlsx`

**Phase 14 Quality Gate**: Cell merging works. All Phase 1-14 tests pass.

---

## Phase 15: UpdateCellCommand

### 15.1 UpdateCellCommand — Custom Cell Processing
**File**: `updatecell.go`, `updatecell_test.go`

**Implementation**:
```go
type CellDataUpdater interface {
    UpdateCellData(cellData *CellData, targetCell CellRef, ctx *Context)
}

type UpdateCellCommand struct {
    BaseCommand
    Updater string  // context key for CellDataUpdater
}
```

**Tests**:
- `TestUpdateCellCommand_BasicUpdate` — updater modifies cell
- `TestUpdateCellCommand_FormulaUpdate` — updater modifies formula (SUM range)
- `TestUpdateCellCommand_NilUpdater` — missing updater → error

**Test template**: `testdata/updatecell_basic.xlsx`

**Phase 15 Quality Gate**: UpdateCell works. All Phase 1-15 tests pass.

---

## Phase 16: Nested Commands & Complex Scenarios

### 16.1 Nested Each
**Tests** (parity with JXLS `NestedSumsTest`):
- `TestNested_EachInsideEach` — departments with employees
- `TestNested_EachInsideEach_WithSums` — nested with SUM formulas
- `TestNested_EachInsideEach_WithIf` — nested with conditional inside

**Test template**: `testdata/nested_each.xlsx`, `testdata/nested_each_sums.xlsx`

### 16.2 Each + If Combined
**Tests** (parity with JXLS `If01Test`):
- `TestCombined_EachWithIf` — conditional rows in loop
- `TestCombined_EachWithIfElse` — if/else inside loop
- `TestCombined_EachWithSelect_vs_If` — select attribute vs jx:if equivalence

### 16.3 Complex Formula Scenarios
**Tests** (parity with JXLS issue-based formula tests):
- `TestComplex_FormulaWithNestedEach` — formulas referencing nested loop cells
- `TestComplex_FormulaWithIf` — formulas with conditional cells
- `TestComplex_CrossSheetFormula` — formula referencing another sheet
- `TestComplex_ExternalFormula` — formula referencing cells outside area

### 16.4 Edge Cases
**Tests** (parity with JXLS issue tests):
- `TestEdge_EmptyListFormula` — formula with empty each → default value
- `TestEdge_SingleRowFormula` — formula with 1-row each
- `TestEdge_LargeDataSet` — 1000+ rows (performance baseline)
- `TestEdge_SpecialCharInData` — data with <, >, &, quotes
- `TestEdge_LongString` — very long cell value
- `TestEdge_UnicodeData` — Unicode characters in data
- `TestEdge_MultipleSheets` — template with multiple sheets
- `TestEdge_ConditionalFormatting` — conditional formatting preserved

**Test templates**: `testdata/nested_*.xlsx`, `testdata/edge_*.xlsx`

**Phase 16 Quality Gate**: Complex scenarios work correctly. All tests pass.

---

## Phase 17: Custom Commands & Extensibility

### 17.1 Custom Command Registration
**File**: `command.go`, integration in `filler.go`

**Implementation**:
- Users can implement `Command` interface
- Register via `WithCommand("name", factoryFunc)` option
- Parser recognizes custom command names in comments

**Tests**:
- `TestCustomCommand_Registration` — register and use custom command
- `TestCustomCommand_InTemplate` — custom command in cell comment
- `TestCustomCommand_WithAttributes` — custom attributes parsed

### 17.2 Custom Expression Notation
**Tests** (parity with JXLS `NotationTest`):
- `TestCustomNotation_Brackets` — `{{...}}` notation
- `TestCustomNotation_TripleBracket` — `[[[...]]]` notation

### 17.3 Pre-Write Actions
**Tests**:
- `TestPreWrite_CustomAction` — user-defined pre-write logic executed
- `TestPreWrite_MultipleActions` — multiple pre-write actions in order

**Phase 17 Quality Gate**: Library is extensible. Custom commands and notations work.

---

## Phase 18: Streaming Support

### 18.1 Streaming for Large Files
**File**: `excelize_tx.go` (streaming mode in transformer)

**Implementation**:
- Use excelize's `StreamWriter` for large datasets
- Only applicable to simple forward-only templates
- Limited formula support (no backward references)

**Tests**:
- `TestStreaming_LargeDataSet` — 10,000 rows with streaming
- `TestStreaming_MemoryUsage` — verify lower memory allocation
- `TestStreaming_FormulasDisabled` — formulas skipped in streaming mode

**Phase 18 Quality Gate**: Large files can be processed without excessive memory.

---

## Phase 19: Performance & Benchmarks

### 19.1 Benchmark Suite
**File**: `benchmark_test.go`

```go
func BenchmarkFill_100Rows(b *testing.B)
func BenchmarkFill_1000Rows(b *testing.B)
func BenchmarkFill_10000Rows(b *testing.B)
func BenchmarkFill_NestedLoops(b *testing.B)
func BenchmarkFill_WithFormulas(b *testing.B)
func BenchmarkFill_WithStreaming(b *testing.B)
func BenchmarkExprEvaluate(b *testing.B)
func BenchmarkParseComment(b *testing.B)
```

**Targets** (JXLS benchmark: 30,000 rows in 5.2s):
- 1,000 rows: < 500ms
- 10,000 rows: < 3s
- 30,000 rows: < 6s (parity with JXLS)

### 19.2 Memory Profiling
- Verify no memory leaks after Fill() returns
- Verify transformer cleanup
- Verify large object release

**Phase 19 Quality Gate**: Performance meets or exceeds JXLS benchmarks. No memory leaks.

---

## Phase 20: Polish & Documentation

### 20.1 Error Messages
- All errors wrapped with context (file, cell, command)
- User-facing errors are clear and actionable
- Internal errors logged, not leaked

### 20.2 API Documentation
- GoDoc comments on all public types and functions
- Examples in `example_test.go`:
  - `ExampleFill` — simplest usage
  - `ExampleNewFiller` — builder usage with options
  - `ExampleFill_withIf` — conditional example
  - `ExampleFill_groupBy` — grouping example

### 20.3 README
- Quick start guide
- Template syntax reference
- Comparison with JXLS
- Migration guide for JXLS users

**Phase 20 Quality Gate**: `go doc` output is clean. Examples compile and run.

---

## Test Parity Matrix: JXLS → goxls

| JXLS Test | goxls Test | Phase |
|-----------|------------|-------|
| **CellRef/AreaRef/Size** | | |
| AreaRefContainsTest (9 cases) | TestAreaRef_Contains_* (10 cases) | 1 |
| SizeTest | TestSize_* | 1 |
| CellDataJTest | TestCellData_* | 1 |
| **Expression** | | |
| JexlExpressionEvaluatorTest | TestExpr_* (18 cases) | 2 |
| JexlExpressionEvaluatorNoThreadLocalTest | TestExpr_ConcurrencySafe | 2 |
| ExpressionEvaluatorTest | TestExpr_Error* | 2 |
| NotationTest | TestCustomNotation_* | 17 |
| **Context** | | |
| ContextTest | TestContext_* | 3 |
| RunVar behavior | TestContext_RunVar* | 3 |
| **Transformer** | | |
| PoiTransformer (cell ops) | TestTransformer_* | 4 |
| LastCommentedCellTest | TestTransformer_GetCommentedCells | 4 |
| **Parser** | | |
| LiteralsExtractorTest | TestParseComment_MultiLine | 5 |
| XlsCommentAreaBuilder | TestParseComment_*, TestBuildAreas_* | 5 |
| **Area** | | |
| XlsAreaTest | TestArea_* | 6 |
| **Each** | | |
| EachTest (basic) | TestEachCommand_Basic* | 7 |
| EachCommandTest (unit) | TestEachCommand_EmptyList, NilList | 7 |
| EachTest.varIndex | TestEachCommand_VarIndex | 7 |
| DirectionRightTest | TestEachCommand_DirectionRight_* | 10 |
| SelectTest | TestEachCommand_Select_* | 10 |
| EachTest.orderBy | TestEachCommand_OrderBy_* | 10 |
| OrderByComparatorTest | TestOrderByComparator_* | 10 |
| GroupByTest (5 variants) | TestEachCommand_GroupBy_* | 10 |
| GroupOrderTest | TestEachCommand_GroupBy_* | 10 |
| MultiSheetTest (5 scenarios) | TestEachCommand_MultiSheet_* | 10 |
| PoiSafeSheetNameBuilderUnitTest | TestSafeSheetName | 10 |
| ScalarsTest | TestEachCommand_*Slice | 10 |
| **If** | | |
| IfTest (3 variants) | TestIfCommand_* | 8 |
| If01Test (multilingual) | TestCombined_EachWithIf* | 16 |
| **Formula** | | |
| CreateTargetCellRefTest (12 cases) | TestFormulaProcessor_* | 11 |
| FormulaProcessorsTest (3 modes) | TestFormulaProcessor_*, TestFastFormulaProcessor_* | 11 |
| FastFormulaProcessorTest | TestFastFormulaProcessor_* | 11 |
| **Grid** | | |
| GridTest (3 variants) | TestGridCommand_* | 12 |
| **Image** | | |
| ImageTest | TestImageCommand_* | 13 |
| **MergeCells** | | |
| AreaColumnMergeTest | TestMergeCellsCommand_* | 14 |
| **UpdateCell** | | |
| UpdateCell (via template tests) | TestUpdateCellCommand_* | 15 |
| **Nested/Complex** | | |
| NestedSumsTest (3 variants) | TestNested_* | 16 |
| IssueB116Test (formulas) | TestComplex_Formula* | 16 |
| **Integration** | | |
| ExceptionHandlerTest | Error handling in all phases | 9+ |
| ClearTemplateCellsTest | TestFill_ClearsTemplateCells | 9 |
| PreWriteTest | TestPreWrite_* | 17 |
| SubtotalTest (command extension) | TestCustomCommand_* | 17 |
| YellowCommandTest (custom cmd) | TestCustomCommand_* | 17 |
| ConditionalFormattingTest | TestEdge_ConditionalFormatting | 16 |
| DataValidationTest | TestEdge_DataValidation | 16 |
| **Excluded (by design)** | | |
| DatabaseAccess / JDBC tests | N/A — excluded, anti-pattern | - |
| DynaBeanTest | N/A — Java-specific concept | - |
| JSR310Test | N/A — Go uses time.Time natively | - |
| JexlContextFactoryTest | N/A — Go uses expr-lang | - |
| IssueSxssfTransformerTest | Covered by streaming tests | 18 |

---

## Quality Enforcement Rules

### Every Phase Must:
1. **Write tests FIRST** (red) → implement (green) → refactor (clean)
2. Pass `go test ./... -race -count=1` (race detector enabled)
3. Pass `go vet ./...`
4. Maintain ≥80% test coverage (`go test -coverprofile=cover.out`)
5. Pass ALL prior phase tests (regression gate)
6. Have no `TODO` or `FIXME` left unaddressed in committed code

### Test File Conventions:
- Unit tests: `*_test.go` next to source file
- Integration tests: `goxls_test.go` (uses test templates from `testdata/`)
- Benchmark tests: `benchmark_test.go`
- Test data: `testdata/*.xlsx` (committed to repo)
- Test helpers: `testutil_test.go` (assertion helpers, template builders)

### Commit Convention:
- One commit per sub-phase (e.g., "Phase 1.1: CellRef implementation")
- Commit message includes test count: "Phase 1.1: CellRef (12 tests)"
- No commits with failing tests

---

## Dependency Lock

```
github.com/xuri/excelize/v2    v2.9.0+   # Excel I/O
github.com/expr-lang/expr      v1.16+    # Expression evaluation
github.com/stretchr/testify    v1.9+     # Test assertions (test only)
```

No other dependencies. Standard library for everything else.

---

## Summary: Phase → Feature → Test Count (Estimated)

| Phase | Feature | Est. Tests |
|-------|---------|------------|
| 0 | Bootstrap | 0 |
| 1 | CellRef, AreaRef, Size, CellData | ~25 |
| 2 | Expression evaluator + parser | ~20 |
| 3 | Context + RunVar | ~12 |
| 4 | Transformer (excelize) | ~18 |
| 5 | Template parser (comments → commands) | ~22 |
| 6 | Area processing engine | ~12 |
| 7 | EachCommand (basic DOWN) | ~13 |
| 8 | IfCommand | ~6 |
| 9 | End-to-end Filler integration | ~16 |
| 10 | Advanced each (RIGHT, select, orderBy, groupBy, multisheet) | ~30 |
| 11 | Formula processing | ~16 |
| 12 | GridCommand | ~6 |
| 13 | ImageCommand | ~4 |
| 14 | MergeCellsCommand | ~4 |
| 15 | UpdateCellCommand | ~3 |
| 16 | Nested commands + complex scenarios | ~12 |
| 17 | Custom commands + extensibility | ~5 |
| 18 | Streaming | ~3 |
| 19 | Benchmarks | ~8 |
| 20 | Polish | ~4 |
| **Total** | | **~239** |
