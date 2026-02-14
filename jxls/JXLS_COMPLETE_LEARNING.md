# JXLS 3.0 Complete Learning Guide

## Table of Contents
1. [What is JXLS?](#1-what-is-jxls)
2. [Core Philosophy](#2-core-philosophy)
3. [Template Grammar & Markup](#3-template-grammar--markup)
4. [Commands Reference](#4-commands-reference)
5. [Expression System](#5-expression-system)
6. [Architecture & Source Code](#6-architecture--source-code)
7. [Builder API](#7-builder-api)
8. [Formula Processing](#8-formula-processing)
9. [Streaming & Optimization](#9-streaming--optimization)
10. [Database Access](#10-database-access)
11. [Key Design Patterns](#11-key-design-patterns)
12. [Source Code Map](#12-source-code-map)

---

## 1. What is JXLS?

JXLS is a Java 17 library (v3.0+) that generates Excel reports from templates. Instead of building
Excel files programmatically cell-by-cell, you create an Excel template with special markup
(directives in cell comments + expressions in cell values), then JXLS fills the template with data.

**Analogy**: Like email templates where CSS handles design and template markup handles data population.
Business people design beautiful Excel files in Excel editors, developers annotate them with JXLS
markup, and the library fills them with data at runtime.

**License**: Apache 2.0
**Repository**: https://github.com/jxlsteam/jxls
**Depends on**: Apache POI (for Excel I/O), Apache Commons JEXL (for expression evaluation)

### Basic Usage

```java
// 1. Prepare data
Map<String, Object> data = new HashMap<>();
data.put("employees", employeeList);

// 2. Process template
JxlsPoiTemplateFillerBuilder.newInstance()
    .withTemplate("template.xlsx")
    .buildAndFill(data, new JxlsOutputFile(new File("report.xlsx")));
```

---

## 2. Core Philosophy

1. **Template-first**: Design is done in Excel by business users. No coding for cosmetics.
2. **Separation of concerns**: Design/formatting in Excel, data population in code.
3. **Markup in comments**: Template directives (commands) go into Excel cell comments/notes.
4. **Expressions in cells**: Data binding expressions go directly into cell values using `${...}`.
5. **Working area**: A `jx:area` command demarcates the region JXLS processes.
6. **Commands drive transformation**: Commands (each, if, grid, etc.) control how areas expand.
7. **Formatting preservation**: Cell styles, fonts, borders, colors from the template are preserved.
8. **Comments removed in output**: All JXLS command comments are removed from the final output.

---

## 3. Template Grammar & Markup

### 3.1 Supported Formats
- `.xlsx` (preferred)
- `.xls` (legacy)
- `.xlsm` (macro-enabled)

### 3.2 Two Components of a Template

**Component 1: Commands in Cell Comments (Notes)**
- Created via right-click > "New Note" in Excel
- Syntax: `jx:COMMAND_NAME(attr1="value1" attr2="value2" ... lastCell="CELL_REF")`
- Multiple commands can be in one comment (one per line)
- Comments are REMOVED from output

**Component 2: Expressions in Cell Values**
- Default syntax: `${expression}`
- Configurable notation (e.g., `{{expression}}`)
- Uses JEXL (Java Expression Language) under the hood
- Property access via dot notation: `${employee.name}`

### 3.3 Working Area Demarcation

Every template needs a `jx:area` command to define the processing region:

```
Cell A1 comment: jx:area(lastCell="D4")
```

This tells JXLS: "Process everything from A1 to D4."

### 3.4 Template Structure Example

```
┌─────────────────────────────────────────────────────────┐
│ A1 [comment: jx:area(lastCell="C2")]                    │
│                                                         │
│    A1           B1              C1                       │
│ ┌──────────┬────────────┬──────────────┐                │
│ │ Name     │ Birth Date │ Payment      │  ← Row 1: Header│
│ ├──────────┼────────────┼──────────────┤                │
│ │${e.name} │${e.birth}  │${e.payment}  │  ← Row 2: Data │
│ └──────────┴────────────┴──────────────┘                │
│                                                         │
│ A2 [comment: jx:each(items="employees" var="e"          │
│              lastCell="C2")]                             │
└─────────────────────────────────────────────────────────┘
```

**Key Points**:
- `jx:area` at A1 defines entire processing region (A1:C2)
- `jx:each` at A2 iterates over `employees`, expanding rows downward
- `${e.name}`, `${e.birth}`, `${e.payment}` are data-binding expressions
- All formatting (bold headers, date formats, number formats) comes from the template

---

## 4. Commands Reference

### 4.1 Command Syntax

```
jx:COMMAND_NAME(attr1="value1" attr2="value2" lastCell="CELL_REF")
```

- Commands go in Excel cell comments/notes
- The cell containing the comment is the **start cell** (top-left corner)
- `lastCell` defines the **end cell** (bottom-right corner)
- Together they define the **command area** (rectangular region)

### 4.2 jx:area - Working Area

**Purpose**: Marks the worksheet region for JXLS processing. Required. Usually at cell A1.

```
jx:area(lastCell="F10")
```

| Attribute | Required | Description |
|-----------|----------|-------------|
| lastCell  | Yes      | Bottom-right cell of the processing area |

### 4.3 jx:each - Iteration (Most Important Command)

**Purpose**: Iterates over collections, creating rows/columns/sheets.

```
jx:each(items="employees" var="e" lastCell="C2")
```

| Attribute | Required | Description |
|-----------|----------|-------------|
| items     | Yes      | Expression returning Iterable or array |
| var       | Yes      | Variable name for current item |
| lastCell  | Yes      | End of template area to repeat |
| direction | No       | "DOWN" (default, adds rows) or "RIGHT" (adds columns) |
| varIndex  | No       | Variable name for 0-based iteration index |
| select    | No       | Filter expression (e.g., `"e.payment > 2000"`) |
| groupBy   | No       | Property to group by (prepend var+".") |
| groupOrder| No       | "ASC" or "DESC" for group sorting |
| orderBy   | No       | Comma-separated sort properties with ASC/DESC |
| multisheet| No       | Variable containing sheet names list |
| cellRefGenerator | No | Custom cell reference strategy |
| oldSelectBehavior | No | "true" for v2.12 select behavior |

**Direction Examples**:
```
# Vertical (default) - adds rows
jx:each(items="employees" var="e" direction="DOWN" lastCell="C2")

# Horizontal - adds columns
jx:each(items="months" var="m" direction="RIGHT" lastCell="A2")
```

**Filtering**:
```
jx:each(items="employees" var="e" select="e.payment > 2000" lastCell="C2")
```

**Sorting**:
```
jx:each(items="employees" var="e" orderBy="e.name ASC, e.payment DESC" lastCell="C2")
```
Sort modifiers: `ASC`, `DESC`, `ASC_ignoreCase`, `DESC_ignoreCase`

**Grouping**:
```
jx:each(items="employees" var="g" groupBy="g.department" groupOrder="ASC" lastCell="C4")
```
Access: `${g.item.department}` (group key), iterate `g.items` (grouped items)

**Multi-sheet**:
```
jx:each(items="departments" var="dept" multisheet="sheetNames" lastCell="C10")
```
Each iteration creates a new worksheet.

**Iteration Index**:
```
jx:each(items="employees" var="e" varIndex="idx" lastCell="C2")
```
Then use `${idx}` in cells (0-based).

### 4.4 jx:if - Conditional Rendering

**Purpose**: Show/hide areas based on conditions.

```
jx:if(condition="e.payment > 2000" lastCell="C2")
```

With else area:
```
jx:if(condition="e.payment > 2000" lastCell="C2" areas=["A2:C2","A3:C3"])
```

| Attribute | Required | Description |
|-----------|----------|-------------|
| condition | Yes      | Boolean expression |
| lastCell  | Yes      | End of "if" area |
| areas     | No       | Two area refs: [ifArea, elseArea] |

- Null conditions treated as `false` (v3+)
- For filtering inside `jx:each` without else, prefer `select` attribute on `jx:each`

### 4.5 jx:grid - Dynamic Grid

**Purpose**: Create grids with dynamic columns and rows.

```
jx:grid(headers="headers" data="items" areas=["A3:A3","A4:A4"] lastCell="A4")
```

| Attribute   | Required | Description |
|-------------|----------|-------------|
| headers     | Yes      | Iterable of header values |
| data        | Yes      | Iterable of data rows |
| lastCell    | Yes      | End of grid area |
| areas       | Yes      | [headerArea, bodyArea] |
| props       | No       | Comma-separated property names (for object data) |
| formatCells | No       | Type-to-cell format mapping (e.g., "Double:B4,Date:C4") |

**Data can be**:
- `List<List<Object>>` - each inner list is a row
- `List<Object[]>` - each array is a row
- Objects with `props` - library extracts properties

### 4.6 jx:image - Image Embedding (POI-specific)

```
jx:image(src="imageVar" imageType="PNG" lastCell="A2")
```

| Attribute | Required | Description |
|-----------|----------|-------------|
| src       | Yes      | Expression returning byte[] |
| lastCell  | Yes      | Area for image placement |
| imageType | No       | PNG (default), JPEG, EMF, WMF, PICT, DIB |
| scaleX    | No       | Width scaling factor (double) |
| scaleY    | No       | Height scaling factor (double) |

Image must be pre-loaded as byte array:
```java
byte[] imageBytes = ImageCommand.toByteArray(inputStream);
data.put("imageVar", imageBytes);
```

### 4.7 jx:mergeCells - Cell Merging (POI-specific)

```
jx:mergeCells(lastCell="D2" cols="4" rows="2")
```

| Attribute | Required | Description |
|-----------|----------|-------------|
| lastCell  | Yes      | Merge range endpoint |
| cols      | No       | Number of columns to merge |
| rows      | No       | Number of rows to merge |
| minCols   | No       | Minimum columns for merge to occur |
| minRows   | No       | Minimum rows for merge to occur |

**Restriction**: Cannot be applied to already-merged cells.

### 4.8 jx:updateCell - Custom Cell Processing

```
jx:updateCell(lastCell="E4" updater="totalCellUpdater")
```

| Attribute | Required | Description |
|-----------|----------|-------------|
| updater   | Yes      | Context key for CellDataUpdater implementation |
| lastCell  | Yes      | Cell area |

The updater modifies cell data BEFORE transformation:
```java
class TotalCellUpdater implements CellDataUpdater {
    public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
        if (cellData.isFormulaCell() && cellData.getFormula().equals("SUM(E2)")) {
            String formula = String.format("SUM(E2:E%d)", targetCell.getRow());
            cellData.setEvaluationResult(formula);
        }
    }
}
```

### 4.9 jx:params - Cell-Level Parameters

**Purpose**: Single-cell configuration. No lastCell, no command class.

```
jx:params(defaultValue="1")
jx:params(formulaStrategy="BY_COLUMN")
```

| Attribute       | Description |
|-----------------|-------------|
| defaultValue    | Default value when referenced cells are removed (default: "=0") |
| formulaStrategy | "BY_COLUMN" restricts formula refs to same column |

---

## 5. Expression System

### 5.1 Default Syntax

```
${expression}
```

### 5.2 JEXL 3.3 (Default Expression Language)

Based on Apache Commons JEXL:

| Feature | Example |
|---------|---------|
| Property access | `${obj.propertyName}` |
| Nested property | `${obj.address.city}` |
| Arithmetic | `${obj.num1 + obj.num2}` |
| Comparison | `${obj.num1 > obj.num2}` |
| Logical | `${obj.num1 > 0 && obj.num2 < 100}` |
| Ternary | `${obj.flag ? "Yes" : "No"}` |
| Null-safe | `${obj.name??"default"}` |
| Empty check | `${empty(list)}` |
| Size check | `${size(list)}` |
| String concat | `${"Hello " + obj.name}` |

### 5.3 Custom Notation

```java
builder.withExpressionNotation("{{", "}}")
```
Then templates use `{{expression}}` instead of `${expression}`.

### 5.4 Alternative Expression Engines

- JSR 223 scripting engines (e.g., Spring Expression Language)
- Custom implementations via `ExpressionEvaluatorFactory` interface

---

## 6. Architecture & Source Code

### 6.1 Module Structure

```
jxls-project/
├── jxls/          # Core module (template engine logic)
│   └── src/main/java/org/jxls/
│       ├── command/       # Command interface + EachCommand, IfCommand, GridCommand, etc.
│       ├── area/          # Area, XlsArea, CommandData
│       ├── common/        # CellRef, AreaRef, Size, CellData, Context, etc.
│       ├── expression/    # ExpressionEvaluator, JexlExpressionEvaluator
│       ├── transform/     # Transformer interface, AbstractTransformer
│       ├── formula/       # FormulaProcessor, StandardFormulaProcessor
│       └── builder/       # JxlsTemplateFillerBuilder, JxlsTemplateFiller, XlsCommentAreaBuilder
│
└── jxls-poi/      # POI implementation module
    └── src/main/java/org/jxls/
        ├── command/       # ImageCommand, MergeCellsCommand, etc.
        └── transform/poi/ # PoiTransformer, PoiCellData, JxlsPoiTemplateFillerBuilder
```

### 6.2 Processing Pipeline

```
User Code
    │
    ▼
JxlsPoiTemplateFillerBuilder (configure options)
    │
    ▼
JxlsTemplateFiller.fill(data, output)
    │
    ├── 1. createTransformer()
    │       └── PoiTransformerFactory opens template workbook via Apache POI
    │       └── Reads all sheets, rows, cells, comments into SheetData/RowData/CellData
    │
    ├── 2. installCommands() → XlsCommentAreaBuilder
    │       └── Scans all cell comments for "jx:" prefix
    │       └── Parses command syntax into Command objects
    │       └── Builds Area hierarchy (XlsArea with embedded CommandData)
    │       └── Returns List<Area> (one per jx:area directive)
    │
    ├── 3. processAreas(data)
    │       └── Creates Context with data map
    │       └── For each Area:
    │           └── Area.applyAt(startCell, context)
    │               └── For each cell in area:
    │                   └── If cell has command: command.applyAt()
    │                   │   ├── EachCommand: iterates items, sets var, calls sub-area.applyAt()
    │                   │   ├── IfCommand: evaluates condition, applies ifArea or elseArea
    │                   │   └── etc.
    │                   └── If cell has expression: evaluate ${...} and set result
    │                   └── Transformer.transform(): copy cell value/style to target position
    │
    ├── 4. processFormulas()
    │       └── StandardFormulaProcessor updates all formula cell references
    │       └── Maps source cell refs to target cell refs after expansion
    │
    ├── 5. preWrite() → PreWriteAction hooks
    │
    └── 6. write() → PoiTransformer writes workbook to output
```

### 6.3 Key Interfaces

**Command** - Contract for all template commands:
```java
interface Command {
    String getName();
    List<Area> getAreaList();
    Command addArea(Area area);
    Size applyAt(CellRef cellRef, Context context);  // Main execution
    void reset();
    void setShiftMode(String mode);  // "inner" or "adjacent"
}
```

**Area** - Rectangular region in a sheet:
```java
interface Area {
    Size applyAt(CellRef cellRef, Context context);
    CellRef getStartCellRef();
    Size getSize();
    List<CommandData> getCommandDataList();
    void addCommand(AreaRef ref, Command command);
    Transformer getTransformer();
    void processFormulas(FormulaProcessor fp);
}
```

**Transformer** - Excel I/O abstraction:
```java
interface Transformer {
    void transform(CellRef src, CellRef target, Context context, boolean updateRowHeight);
    void write();
    CellData getCellData(CellRef cellRef);
    List<CellData> getCommentedCells();  // For template parsing
    void setFormula(CellRef cellRef, String formula);
    Set<CellData> getFormulaCells();
    void clearCell(CellRef cellRef);
    boolean deleteSheet(String name);
    // ... many more
}
```

**ExpressionEvaluator** - Expression evaluation:
```java
interface ExpressionEvaluator {
    Object evaluate(String expression, Map<String, Object> data);
    Object evaluate(Map<String, Object> data);  // Pre-compiled
    String getExpression();
    boolean isConditionTrue(Context context);
}
```

**Context** (internal, extends PublicContext):
```java
interface PublicContext {
    Object getVar(String name);
    Object getRunVar(String name);
    void putVar(String name, Object value);
    void removeVar(String name);
    boolean containsVar(String name);
    Object evaluate(String expression);
    boolean isConditionTrue(String condition);
}
```

### 6.4 Key Data Structures

**CellRef** - Single cell reference:
- `sheetName`, `row` (0-based), `col` (0-based)
- Can parse from string: `new CellRef("Sheet1!A1")`

**AreaRef** - Rectangular area:
- `firstCellRef`, `lastCellRef`
- Can parse from string: `new AreaRef("A1:C5")`

**Size** - Width and height:
- `width` (columns), `height` (rows)

**CellData** - Everything about a cell:
- `cellRef`, `cellValue`, `cellType`, `cellComment`
- `formula`, `evaluationResult`
- `formulaStrategy` (DEFAULT, BY_COLUMN, BY_ROW)
- `defaultValue` (from jx:params)
- `targetPos` (list of target positions after transformation)
- `CellType` enum: STRING, NUMBER, BOOLEAN, DATE, LOCAL_DATE, LOCAL_TIME, LOCAL_DATETIME, ZONED_DATETIME, INSTANT, FORMULA, BLANK, ERROR

**GroupData** - For groupBy:
- `item` (group key), `items` (collection of grouped items)

**CommandData** - Command + its area:
- `command`, `startCellRef`, `size`

### 6.5 Cell Shift Strategies

When a command expands (e.g., `jx:each` adds rows), other cells must shift:

- **InnerCellShiftStrategy** (default): Cells inside the command area shift
- **AdjacentCellShiftStrategy**: Cells adjacent to the command area also shift

---

## 7. Builder API

### 7.1 Standard Usage

```java
JxlsPoiTemplateFillerBuilder.newInstance()
    .withTemplate("template.xlsx")
    .buildAndFill(data, new JxlsOutputFile(new File("output.xlsx")));
```

### 7.2 All Builder Options

| Method | Default | Description |
|--------|---------|-------------|
| `withTemplate(String/File/URL/InputStream)` | - | Template source |
| `withExpressionEvaluatorFactory(...)` | JexlImpl | Expression engine |
| `withExpressionNotation(begin, end)` | `${`, `}` | Expression delimiters |
| `withLogger(JxlsLogger)` | PoiExceptionLogger | Logging |
| `withExceptionThrower()` | - | Throw exceptions instead of logging |
| `withFormulaProcessor(FormulaProcessor)` | StandardFormulaProcessor | Formula handling |
| `withFastFormulaProcessor()` | - | 10x faster, limited support |
| `withUpdateCellDataArea(boolean)` | true | Track cell references for formulas |
| `withIgnoreColumnProps(boolean)` | true | Skip auto column width |
| `withIgnoreRowProps(boolean)` | true | Skip auto row height |
| `withRecalculateFormulasBeforeSaving(boolean)` | true | Pre-save recalculation |
| `withRecalculateFormulasOnOpening(boolean)` | false | Open-time recalculation |
| `withKeepTemplateSheet(KeepTemplateSheet)` | DELETE | Template sheet: DELETE/HIDE/KEEP |
| `withAreaBuilder(AreaBuilder)` | XlsCommentAreaBuilder | Command parser |
| `withCommand(name, Class)` | - | Register custom commands |
| `withClearTemplateCells(boolean)` | true | Clear unevaluated cells |
| `withStreaming(JxlsStreaming)` | STREAMING_OFF | Streaming config |
| `needsPublicContext(object)` | - | Inject context into objects |
| `withPreWriteAction(action)` | - | Pre-write hooks |
| `withRunVarAccess(access)` | - | Custom loop variable storage |
| `withTransformerFactory(factory)` | PoiTransformerFactory | Transformer creation |

### 7.3 Output Options

```java
// File output
filler.fill(data, new JxlsOutputFile(new File("output.xlsx")));

// Stream output
filler.fill(data, output);

// Byte array output
byte[] bytes = builder.buildAndFill(data);
```

---

## 8. Formula Processing

### 8.1 How It Works

1. Template contains formulas like `=SUM(C2:C2)`
2. During `jx:each` expansion, C2 might become C2:C10 (10 employees)
3. FormulaProcessor updates the formula to `=SUM(C2:C10)`
4. This happens AFTER all command processing, BEFORE writing

### 8.2 Processors

| Processor | Speed | Capabilities |
|-----------|-------|--------------|
| StandardFormulaProcessor | Normal | Full template support, all formula types |
| FastFormulaProcessor | 10x faster | Limited, simple templates only |
| null (disabled) | Fastest | No formula processing |

### 8.3 Formula Strategy (via jx:params)

- **DEFAULT**: References expand to all target cells
- **BY_COLUMN**: Only reference cells in the same column
- **BY_ROW**: Only reference cells in the same row

### 8.4 Default Values (via jx:params)

When formula references are removed (e.g., empty `jx:each`), the formula gets a default:
- Default: `=0`
- Custom: `jx:params(defaultValue="1")`

---

## 9. Streaming & Optimization

### 9.1 Streaming Modes

| Mode | Description |
|------|-------------|
| STREAMING_OFF | Default. Full workbook in memory. |
| STREAMING_ON | All sheets stream. Rows written to disk immediately. |
| AUTO_DETECT | Sheets with `sheetStreaming="true"` comment stream. |
| streamingWithGivenSheets | Named sheets stream. |

**Limitation**: Streaming cannot forward-reference rows (formula in row 1 can't ref row 2).

### 9.2 Streaming Options

```java
.withStreaming(JxlsStreaming.STREAMING_ON)
    .withOptions(rowAccessWindowSize, compressTmpFiles, useSharedStringsTable)
```

### 9.3 Optimization Priority

1. Enable streaming (most impactful)
2. Optimize sheet order (streaming sheets first)
3. Use FastFormulaProcessor if possible
4. Disable formula recalculation if not needed
5. Avoid inner `jx:if` (use `select` attribute or conditional formatting)
6. Free large objects after use

**Benchmark**: 30,000 rows in 5.2 seconds.

---

## 10. Database Access

```java
Connection conn = ...; // JDBC connection
DatabaseAccess db = new DatabaseAccess(conn);
data.put("jdbc", db);
```

Template:
```
jx:each(items="jdbc.query('select * from employee where payment > ?', 2000)"
        var="emp" lastCell="C4")
```

- Returns `List<Map<String, Object>>`
- Column access: `${emp.name}`, `${emp.payment}`
- Supports parameterized queries with `?`
- Single quotes escaped with backslash: `name=\'Elsa\'`

---

## 11. Key Design Patterns

| Pattern | Where Used |
|---------|------------|
| **Builder** | JxlsTemplateFillerBuilder - fluent API configuration |
| **Strategy** | FormulaProcessor, CellShiftStrategy, CellRefGenerator |
| **Factory** | ExpressionEvaluatorFactory, JxlsTransformerFactory |
| **Template Method** | AbstractCommand, AbstractTransformer |
| **Observer** | AreaListener for transformation events |
| **Context** | Context/ContextImpl for variable management |
| **Adapter** | PoiCellData, PoiSheetData wrapping POI objects |
| **Try-with-resources** | RunVar for automatic variable save/restore |

---

## 12. Source Code Map

### Core Module: `jxls/source/jxls/src/main/java/org/jxls/`

```
command/
  Command.java              # Interface - contract for all commands
  AbstractCommand.java      # Base implementation
  EachCommand.java          # jx:each - iteration (most complex, ~300 lines)
  IfCommand.java            # jx:if - conditional
  GridCommand.java          # jx:grid - dynamic grids
  UpdateCellCommand.java    # jx:updateCell - custom cell processing
  CellDataUpdater.java      # Interface for updateCell
  RunVar.java               # Loop variable management (AutoCloseable)
  CellRefGenerator.java     # Interface for custom cell ref generation
  DynamicSheetNameGenerator.java
  SheetNameGenerator.java

area/
  Area.java                 # Interface - rectangular processing region
  XlsArea.java              # Main implementation (~500 lines)
  CommandData.java           # Command + area binding

common/
  CellRef.java              # Cell reference (sheet, row, col)
  AreaRef.java              # Area reference (two CellRefs)
  Size.java                 # Width x Height
  CellData.java             # Full cell info (~400 lines, expressions, formulas)
  SheetData.java            # Sheet data holder
  RowData.java              # Row data holder
  GroupData.java            # Grouping data (key + items)
  EvaluationResult.java     # Expression eval result
  Context.java              # Internal context interface
  PublicContext.java         # Public context interface
  ContextImpl.java          # Context implementation
  NeedsPublicContext.java   # Interface for context injection

expression/
  ExpressionEvaluator.java           # Interface
  ExpressionEvaluatorFactory.java    # Factory interface
  JexlExpressionEvaluator.java       # JEXL implementation
  JexlExpressionEvaluatorNoThreadLocal.java

transform/
  Transformer.java          # Interface - Excel I/O abstraction
  AbstractTransformer.java  # Base implementation
  ExpressionEvaluatorContext.java    # Expression notation & evaluation
  JxlsTransformerFactory.java       # Factory interface
  TemplateProcessor.java    # Template preprocessor interface

formula/
  FormulaProcessor.java              # Interface
  AbstractFormulaProcessor.java      # Base implementation
  StandardFormulaProcessor.java      # Full formula processing
  FastFormulaProcessor.java          # Optimized formula processing

builder/
  JxlsTemplateFillerBuilder.java     # Main builder (~400 lines)
  JxlsTemplateFiller.java            # Template execution engine
  JxlsOptions.java                   # Options transport
  AreaBuilder.java                   # Interface for area parsing
  xls/
    XlsCommentAreaBuilder.java       # Parses cell comments for commands
```

### POI Module: `jxls/source/jxls-poi/src/main/java/org/jxls/`

```
command/
  ImageCommand.java             # jx:image
  MergeCellsCommand.java        # jx:mergeCells
  AbstractMergeCellsCommand.java
  AutoRowHeightCommand.java     # Auto row height
  AreaColumnMergeCommand.java   # Column merge in areas

transform/poi/
  PoiTransformer.java                    # POI Transformer implementation (~800 lines)
  PoiTransformerFactory.java             # Creates PoiTransformer
  JxlsPoiTemplateFillerBuilder.java      # POI-configured builder
  PoiCellData.java                       # POI cell wrapper
  PoiSheetData.java                      # POI sheet wrapper
  PoiRowData.java                        # POI row wrapper
  PoiUtil.java                           # Utility functions
  PoiSafeSheetNameBuilder.java           # Safe sheet naming
  PoiConditionalFormatting.java          # Conditional formatting
  PoiDataValidations.java               # Data validation
  SelectSheetsForStreamingPoiTransformer.java  # Selective streaming
  WritableCellValue.java                 # Cell value wrapper
  WritableHyperlink.java                 # Hyperlink wrapper
```

### Test Templates: `jxls/source/jxls-poi/src/test/resources/org/jxls/templatebasedtests/`

30+ .xlsx template files covering:
- Basic each/loop (EachTest.xlsx)
- Conditional logic (IfTest.xlsx)
- Nested operations
- Formula processing
- Multi-sheet generation
- Conditional formatting
- Data validation
- Dynamic naming
- Table support

---

## Key Takeaways for Porting

1. **Core is ~15 key classes**: Command hierarchy + Area + Context + Transformer + Expression + FormulaProcessor + Builder
2. **Template parsing is simple**: Scan cell comments for "jx:" prefix, parse attributes with regex
3. **Expression evaluation**: Replace JEXL with Go template expressions or a Go expression library
4. **Transformer is the Excel abstraction**: Replace POI with excelize in Go
5. **Formula processing is the hardest part**: Tracking cell reference changes during expansion
6. **The builder pattern maps well to Go's functional options pattern**
7. **Commands are self-contained**: Each implements `applyAt()` with clear inputs/outputs
