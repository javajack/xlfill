# goxls - Go Port of JXLS Template Engine

## Architecture & Design Plan

---

## 1. Overview

goxls is a Go library that generates Excel reports from templates. Business users design
formatted Excel templates, developers annotate them with template markup (commands in cell
comments + expressions in cell values), and goxls fills them with data at runtime.

**Goal**: Recreate the JXLS development experience in idiomatic Go.

**Underlying Excel library**: [excelize](https://github.com/qax-os/excelize) (Go's Apache POI equivalent)

**Format support**: .xlsx only (excelize limitation; acceptable for modern use)

---

## 2. How It Works (User Perspective)

### Step 1: Business user designs a template in Excel

```
┌──────────────┬────────────┬──────────────┐
│ Name         │ Department │ Salary       │  ← Formatted headers (bold, colored)
├──────────────┼────────────┼──────────────┤
│ ${e.Name}    │ ${e.Dept}  │ ${e.Salary}  │  ← Expression row
├──────────────┼────────────┼──────────────┤
│ Total        │            │ =SUM(C2:C2)  │  ← Formula (auto-adjusted)
└──────────────┴────────────┴──────────────┘
```

Cell A1 comment: `jx:area(lastCell="C3")`
Cell A2 comment: `jx:each(items="employees" var="e" lastCell="C2")`

### Step 2: Developer writes Go code

```go
data := map[string]interface{}{
    "employees": []Employee{
        {Name: "Alice", Dept: "Engineering", Salary: 90000},
        {Name: "Bob", Dept: "Marketing", Salary: 75000},
        {Name: "Charlie", Dept: "Engineering", Salary: 85000},
    },
}

err := goxls.Fill("template.xlsx", "output.xlsx", data)
```

### Step 3: Output

```
┌──────────────┬──────────────┬──────────────┐
│ Name         │ Department   │ Salary       │  ← Original formatting preserved
├──────────────┼──────────────┼──────────────┤
│ Alice        │ Engineering  │ 90000        │
│ Bob          │ Marketing    │ 75000        │
│ Charlie      │ Engineering  │ 85000        │
├──────────────┼──────────────┼──────────────┤
│ Total        │              │ =SUM(C2:C4)  │  ← Formula auto-adjusted!
└──────────────┴──────────────┴──────────────┘
```

---

## 3. Technology Choices

### 3.1 Excel I/O: excelize

**Why excelize?**
- 20K+ GitHub stars, actively maintained (last commit Jan 2026)
- Pure Go, no C dependencies
- Can open existing .xlsx files
- Can read cell comments/notes (critical for template parsing)
- Supports row insertion/deletion
- Supports cell merging, images, formulas
- BSD-3-Clause license

**Key Concern**: excelize doesn't auto-preserve cell styles when setting values.
**Mitigation**: goxls will cache styles before modification and re-apply them.

### 3.2 Expression Evaluation

Options for replacing Apache Commons JEXL:

| Option | Pros | Cons |
|--------|------|------|
| **Go `text/template`** | Built-in, well-known | Different syntax, limited for cell-level use |
| **expr (github.com/expr-lang/expr)** | Fast, safe, Go-native | Different syntax than JEXL |
| **govaluate** | Simple expression eval | Less maintained |
| **Custom parser** | Full control, JXLS-compatible syntax | More work |

**Recommendation**: Use `expr` library with a thin wrapper to support `${...}` notation.
The expressions inside `${...}` will use Go-friendly syntax:

```
JXLS:  ${employee.name}       → goxls: ${employee.Name}     (capitalized for Go exports)
JXLS:  ${e.payment > 2000}    → goxls: ${e.Payment > 2000}
JXLS:  ${empty(list)}         → goxls: ${len(list) == 0}
```

### 3.3 Property Access

In Java, JXLS uses Apache Commons BeanUtils for property access on POJOs.
In Go, we use reflection on structs, or map[string]interface{} access.

Data can be:
- `map[string]interface{}` (like Java's Map)
- Structs with exported fields
- Slices/arrays for iteration

---

## 4. Package Structure

```
goxls/
├── go.mod
├── go.sum
├── goxls.go                 # Public API entry point (Fill, NewFiller)
│
├── cmd/                     # Command implementations
│   ├── command.go           # Command interface
│   ├── each.go              # jx:each - iteration
│   ├── condition.go         # jx:if - conditional
│   ├── grid.go              # jx:grid - dynamic grid
│   ├── image.go             # jx:image - image embedding
│   ├── mergecells.go        # jx:mergeCells - cell merging
│   ├── updatecell.go        # jx:updateCell - custom cell processing
│   └── params.go            # jx:params - cell parameters
│
├── area/                    # Area management
│   ├── area.go              # Area interface + XlsArea implementation
│   └── commanddata.go       # Command + area binding
│
├── context/                 # Data context
│   └── context.go           # Context for variable management
│
├── expr/                    # Expression evaluation
│   ├── evaluator.go         # ExpressionEvaluator interface + implementation
│   └── parser.go            # Parse ${...} expressions from cell values
│
├── transform/               # Excel I/O abstraction
│   ├── transformer.go       # Transformer interface
│   ├── excelize.go          # excelize-based implementation
│   ├── celldata.go          # CellData, CellRef, AreaRef, Size
│   └── sheetdata.go         # SheetData, RowData
│
├── formula/                 # Formula processing
│   ├── processor.go         # FormulaProcessor interface
│   └── standard.go          # Standard formula ref updating
│
├── parse/                   # Template parsing
│   └── comment.go           # Parse jx: commands from cell comments
│
└── _testdata/               # Test templates
    ├── basic_each.xlsx
    ├── nested_each.xlsx
    ├── if_else.xlsx
    ├── grid.xlsx
    └── formulas.xlsx
```

---

## 5. Core Interfaces

### 5.1 Command Interface

```go
// Command represents a template processing command (jx:each, jx:if, etc.)
type Command interface {
    // Name returns the command identifier (e.g., "each", "if")
    Name() string
    
    // Areas returns the areas managed by this command
    Areas() []*Area
    
    // AddArea adds an area to this command
    AddArea(area *Area)
    
    // ApplyAt applies this command at the given cell position with the given context.
    // Returns the resulting size (width, height) after expansion.
    ApplyAt(cellRef CellRef, ctx *Context) (Size, error)
    
    // Reset resets the command state for reuse
    Reset()
    
    // SetShiftMode sets how adjacent cells shift ("inner" or "adjacent")
    SetShiftMode(mode string)
}
```

### 5.2 Area Interface

```go
// Area represents a rectangular region in a worksheet
type Area interface {
    // ApplyAt processes this area at the given position
    ApplyAt(cellRef CellRef, ctx *Context) (Size, error)
    
    // StartCell returns the top-left cell of this area
    StartCell() CellRef
    
    // GetSize returns the dimensions of this area
    GetSize() Size
    
    // Commands returns embedded commands
    Commands() []*CommandData
    
    // AddCommand adds a command to this area
    AddCommand(areaRef AreaRef, cmd Command)
    
    // Transformer returns the Excel transformer
    Transformer() Transformer
    
    // ProcessFormulas updates formulas after transformation
    ProcessFormulas(fp FormulaProcessor)
    
    // ClearCells clears template cells in the area
    ClearCells()
}
```

### 5.3 Transformer Interface

```go
// Transformer abstracts Excel I/O operations
type Transformer interface {
    // Transform copies a cell from source to target, evaluating expressions
    Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error
    
    // Write writes the workbook to the output
    Write(w io.Writer) error
    
    // GetCellData returns cell information
    GetCellData(ref CellRef) (*CellData, error)
    
    // GetCommentedCells returns all cells with comments (for template parsing)
    GetCommentedCells() ([]*CellData, error)
    
    // SetFormula sets a formula on a cell
    SetFormula(ref CellRef, formula string) error
    
    // ClearCell clears a cell's content
    ClearCell(ref CellRef) error
    
    // InsertRows inserts rows at the given position
    InsertRows(sheet string, row, count int) error
    
    // DeleteSheet removes a sheet
    DeleteSheet(name string) error
    
    // SetHidden hides/shows a sheet
    SetHidden(name string, hidden bool) error
    
    // MergeCells merges a cell range
    MergeCells(sheet string, topLeft, bottomRight string) error
    
    // AddImage inserts an image
    AddImage(sheet string, ref CellRef, imgBytes []byte, imgType string) error
    
    // GetFormulaCells returns all cells containing formulas
    GetFormulaCells() []*CellData
    
    // GetTargetCellRef returns where a source cell was mapped to
    GetTargetCellRef(src CellRef) []CellRef
    
    // ResetTargetCellRefs clears the source->target mapping
    ResetTargetCellRefs()
}
```

### 5.4 ExpressionEvaluator Interface

```go
// ExpressionEvaluator evaluates template expressions
type ExpressionEvaluator interface {
    // Evaluate evaluates an expression with the given data
    Evaluate(expression string, data map[string]interface{}) (interface{}, error)
    
    // IsConditionTrue evaluates a boolean condition
    IsConditionTrue(condition string, data map[string]interface{}) (bool, error)
}
```

### 5.5 Context

```go
// Context holds template data and provides expression evaluation
type Context struct {
    data      map[string]interface{}  // User-provided data
    runVars   map[string]interface{}  // Loop iteration variables  
    evaluator ExpressionEvaluator
    notationBegin string  // default "${"
    notationEnd   string  // default "}"
}

func (c *Context) GetVar(name string) interface{}
func (c *Context) PutVar(name string, value interface{})
func (c *Context) RemoveVar(name string)
func (c *Context) Evaluate(expression string) (interface{}, error)
func (c *Context) IsConditionTrue(condition string) (bool, error)
```

### 5.6 Data Structures

```go
// CellRef represents a cell reference
type CellRef struct {
    Sheet string
    Row   int  // 0-based
    Col   int  // 0-based
}

// AreaRef represents a rectangular area
type AreaRef struct {
    First CellRef
    Last  CellRef
}

// Size represents width and height
type Size struct {
    Width  int  // columns
    Height int  // rows
}

// CellData holds all information about a cell
type CellData struct {
    Ref             CellRef
    Value           interface{}
    Type            CellType      // String, Number, Boolean, Date, Formula, Blank, Error
    Comment         string
    Formula         string
    EvalResult      interface{}
    FormulaStrategy FormulaStrategy  // Default, ByColumn, ByRow
    DefaultValue    string
    TargetPositions []CellRef
    StyleID         int  // Cached style for preservation
}

type CellType int
const (
    CellTypeBlank CellType = iota
    CellTypeString
    CellTypeNumber
    CellTypeBoolean
    CellTypeDate
    CellTypeFormula
    CellTypeError
)

type FormulaStrategy int
const (
    FormulaDefault FormulaStrategy = iota
    FormulaByColumn
    FormulaByRow
)
```

---

## 6. Public API Design

### 6.1 Simple API (Most Common)

```go
// Fill processes a template file and writes the result to an output file
func Fill(templatePath, outputPath string, data map[string]interface{}) error

// FillBytes processes a template and returns the result as bytes
func FillBytes(templatePath string, data map[string]interface{}) ([]byte, error)

// FillReader processes a template from a reader
func FillReader(template io.Reader, output io.Writer, data map[string]interface{}) error
```

### 6.2 Builder API (Advanced)

```go
filler := goxls.NewFiller(
    goxls.WithTemplate("template.xlsx"),
    goxls.WithExpressionNotation("{{", "}}"),
    goxls.WithFormulaProcessor(goxls.FastFormulas),
    goxls.WithStreaming(true),
    goxls.WithCommand("custom", &MyCustomCommand{}),
    goxls.WithPreWrite(func(t Transformer, ctx *Context) error {
        // Custom pre-write logic
        return nil
    }),
)

err := filler.Fill(data, "output.xlsx")
```

### 6.3 Functional Options Pattern

```go
type Option func(*Filler)

func WithTemplate(path string) Option
func WithTemplateReader(r io.Reader) Option
func WithExpressionNotation(begin, end string) Option
func WithFormulaProcessor(fp FormulaProcessor) Option
func WithStreaming(enabled bool) Option
func WithCommand(name string, cmd Command) Option
func WithPreWrite(action func(Transformer, *Context) error) Option
func WithClearTemplateCells(clear bool) Option
func WithKeepTemplateSheet(keep bool) Option
```

---

## 7. Template Parsing

### 7.1 Comment Syntax (Same as JXLS)

```
jx:COMMAND_NAME(attr1="value1" attr2="value2" lastCell="CELL_REF")
```

### 7.2 Parsing Algorithm

```
1. Open template with excelize
2. For each sheet:
   a. Get all cell comments via excelize.GetComments()
   b. For each comment:
      - Check if starts with "jx:"
      - Parse command name
      - Parse attributes (key="value" pairs)
      - Create Command instance
      - Determine command area (comment cell = start, lastCell attr = end)
3. Build Area hierarchy:
   - jx:area defines root areas
   - Other commands are nested within their containing area
4. Return List of root Areas
```

### 7.3 Expression Parsing in Cells

```
1. For each cell value (string type):
   a. Find all ${...} patterns (configurable notation)
   b. Extract expression content
   c. During processing: evaluate with Context data, replace in cell value
```

---

## 8. Command Implementations

### 8.1 jx:each (EachCommand)

```
Input:
  - items: expression returning slice/array
  - var: variable name for current item
  - direction: "DOWN" (default) or "RIGHT"
  - varIndex: variable for iteration index (optional)
  - select: filter expression (optional)
  - orderBy: sort specification (optional)
  - groupBy/groupOrder: grouping (optional)
  - multisheet: sheet names variable (optional)

Algorithm:
  1. Evaluate "items" expression to get collection
  2. Apply "select" filter if specified
  3. Apply "orderBy" sorting if specified
  4. Apply "groupBy" grouping if specified
  5. For each item (or group):
     a. Set var in context (save old value)
     b. Set varIndex if specified
     c. Calculate target cell (based on direction and iteration)
     d. Call sub-area.ApplyAt(targetCell, context)
     e. Accumulate total size
  6. Restore context (remove var)
  7. Return total size
```

### 8.2 jx:if (IfCommand)

```
Algorithm:
  1. Evaluate condition expression
  2. If true: apply ifArea at target cell
  3. If false: apply elseArea at target cell (if exists)
  4. Return resulting size
```

### 8.3 jx:grid (GridCommand)

```
Algorithm:
  1. Evaluate "headers" expression for column headers
  2. Evaluate "data" expression for data rows
  3. Render header area horizontally for each header
  4. For each data row:
     a. Render body area for each column value
  5. Apply formatCells type mapping
  6. Return total grid size
```

---

## 9. Formula Processing

### 9.1 The Problem

Template has: `=SUM(C2:C2)` (one data row in template)
After expanding 10 rows: formula should become `=SUM(C2:C11)`

### 9.2 Algorithm

```
1. During transformation, track: sourceCell → [targetCell1, targetCell2, ...]
2. After all commands execute:
   a. For each formula cell:
      - Parse formula for cell references
      - For each reference, look up target positions
      - Build new formula with expanded references
      - Handle BY_COLUMN/BY_ROW strategy
      - Handle default values for removed refs
   b. Write updated formulas to output
```

### 9.3 Handling Large Formulas

Excel's SUM() supports max 255 arguments. For larger sets:
- Convert `SUM(C2:C300)` to `C2+C3+C4+...+C300` (addition chain)
- Or use sub-SUMs: `SUM(C2:C255)+SUM(C256:C300)`

---

## 10. Style Preservation Strategy

This is the critical challenge with excelize.

### 10.1 Approach

```go
type StyleCache struct {
    styles map[string]int  // "Sheet!A1" → styleID
}

// Before processing a cell:
func (sc *StyleCache) Save(f *excelize.File, sheet, cell string) {
    styleID, _ := f.GetCellStyle(sheet, cell)
    sc.styles[sheet+"!"+cell] = styleID
}

// After setting a cell value:
func (sc *StyleCache) Restore(f *excelize.File, sheet, cell string) {
    if styleID, ok := sc.styles[sheet+"!"+cell]; ok {
        f.SetCellStyle(sheet, cell, cell, styleID)
    }
}
```

### 10.2 Row Insertion Style Copying

When `jx:each` inserts new rows, each new row should inherit the style of the template row:

```go
func copyRowStyle(f *excelize.File, sheet string, srcRow, dstRow int) {
    // For each cell in source row:
    //   1. Get style from source cell
    //   2. Apply style to corresponding destination cell
}
```

---

## 11. Implementation Phases

### Phase 1: Core Infrastructure
- [ ] CellRef, AreaRef, Size data structures
- [ ] Transformer interface + excelize implementation
- [ ] Template parser (comment parsing, expression detection)
- [ ] Context with variable management
- [ ] Expression evaluator (using expr library)
- [ ] Basic Fill() API

### Phase 2: Essential Commands
- [ ] jx:area command
- [ ] jx:each command (basic: items, var, lastCell, direction DOWN)
- [ ] jx:if command (condition, if/else areas)
- [ ] Style preservation during transformation

### Phase 3: Formula Processing
- [ ] Source-to-target cell mapping
- [ ] StandardFormulaProcessor
- [ ] Formula reference expansion
- [ ] Default values for removed refs

### Phase 4: Advanced Each Features
- [ ] direction="RIGHT"
- [ ] varIndex
- [ ] select (filtering)
- [ ] orderBy (sorting)
- [ ] groupBy/groupOrder (grouping)
- [ ] multisheet

### Phase 5: Additional Commands
- [ ] jx:grid command
- [ ] jx:image command
- [ ] jx:mergeCells command
- [ ] jx:updateCell command
- [ ] jx:params (defaultValue, formulaStrategy)

### Phase 6: Advanced Features
- [ ] Streaming for large files
- [ ] Custom commands support
- [ ] Custom expression notation
- [ ] Pre-write actions
- [ ] Builder API with all options

### Phase 7: Polish
- [ ] Comprehensive tests with template files
- [ ] Performance benchmarks
- [ ] Documentation
- [ ] Examples

---

## 12. Dependency Summary

```
goxls
├── github.com/xuri/excelize/v2     # Excel I/O (Apache POI equivalent)
├── github.com/expr-lang/expr       # Expression evaluation (JEXL equivalent)
└── (standard library only beyond these)
```

---

## 13. Key Differences from JXLS

| Aspect | JXLS (Java) | goxls (Go) |
|--------|-------------|------------|
| Excel library | Apache POI | excelize |
| Expression engine | Apache JEXL | expr-lang/expr |
| Property access | BeanUtils (getters) | Reflection (exported fields) |
| Data types | POJOs + Maps | Structs + Maps |
| Format support | .xlsx, .xls, .xlsm | .xlsx only |
| Thread safety | ThreadLocal caching | goroutine-safe by design |
| Builder pattern | Fluent builder | Functional options |
| Streaming | SXSSF (POI) | excelize StreamWriter |
| Error handling | Exceptions | error return values |
| Context | Interface hierarchy | Simple struct |
| Configuration | XML, annotations | Code only (simpler) |

---

## 14. Expression Syntax Mapping

| JXLS (JEXL) | goxls (expr) | Description |
|-------------|--------------|-------------|
| `${e.name}` | `${e.Name}` | Property access (Go exports) |
| `${e.payment > 2000}` | `${e.Payment > 2000}` | Comparison |
| `${e.a + e.b}` | `${e.A + e.B}` | Arithmetic |
| `${e.flag ? "Y" : "N"}` | `${e.Flag ? "Y" : "N"}` | Ternary |
| `${empty(list)}` | `${len(list) == 0}` | Empty check |
| `${size(list)}` | `${len(list)}` | Size check |
| `${e.name??""}` | `${e.Name ?? ""}` | Null-safe (if expr supports) |
