# Go Excel Libraries Comparison

## For goxls (JXLS port) - Apache POI Equivalent Selection

---

## Winner: excelize (github.com/qax-os/excelize)

**Import**: `github.com/xuri/excelize/v2`
**Stars**: 20,300+ | **License**: BSD-3-Clause | **Last commit**: Jan 2026 (active)
**Go version**: 1.24.0+

---

## Feature Matrix for JXLS Porting Requirements

| Requirement | excelize | tealeg/xlsx | unioffice | extrame/xls |
|-------------|----------|-------------|-----------|-------------|
| Open existing .xlsx | Yes | Yes | Yes | Read-only |
| Read cell comments | Yes (GetComments) | Unknown | Unknown | No |
| Write cell values | Yes (SetCellValue) | Yes | Yes | No |
| Preserve formatting | Manual* | Unknown | Unknown | No |
| Insert/delete rows | Yes | Yes | Yes | No |
| Cell merging | Yes | Unknown | Yes | No |
| Images | Yes | Unknown | Yes | No |
| Formulas | Yes | Unknown | Yes | Read-only |
| .xls (old format) | No | No | No | Read-only |
| Actively maintained | Yes | No (discontinued) | Commercial | Limited |
| Streaming API | Yes | No | Unknown | No |

*excelize does NOT auto-preserve styles when setting cell values. Must manually
cache style with GetCellStyle() and re-apply with SetCellStyle() after SetCellValue().

---

## Critical excelize APIs for goxls

### Template Parsing
```go
// Read cell comments (JXLS commands are in comments)
comments, err := f.GetComments("Sheet1")
// Returns []excelize.Comment with Author, Text, Cell fields
```

### Cell Operations
```go
// Read cell value
value, err := f.GetCellValue("Sheet1", "A1")

// Set cell value (various types)
f.SetCellValue("Sheet1", "A1", "Hello")
f.SetCellValue("Sheet1", "B1", 42)
f.SetCellValue("Sheet1", "C1", time.Now())

// Set formula
f.SetCellFormula("Sheet1", "D1", "SUM(B1:B10)")
```

### Style Preservation (CRITICAL)
```go
// MUST cache before modifying
styleID, _ := f.GetCellStyle("Sheet1", "A1")
f.SetCellValue("Sheet1", "A1", newValue)
f.SetCellStyle("Sheet1", "A1", "A1", styleID)
```

### Row Operations
```go
// Insert rows (shifts existing rows down)
f.InsertRows("Sheet1", 3, 5)  // Insert 5 rows at row 3

// Remove row
f.RemoveRow("Sheet1", 3)

// Set row height
f.SetRowHeight("Sheet1", 1, 30)

// Get row height
height, _ := f.GetRowHeight("Sheet1", 1)
```

### Cell Merging
```go
f.MergeCell("Sheet1", "A1", "C1")
f.UnmergeCell("Sheet1", "A1", "C1")
cells, _ := f.GetMergeCells("Sheet1")
```

### Images
```go
f.AddPicture("Sheet1", "A1", &excelize.Picture{
    File: "image.png",
    Format: excelize.GraphicOptions{...},
})
// Or from bytes using AddPictureFromBytes
```

### Sheet Operations
```go
f.NewSheet("Sheet2")
f.DeleteSheet("Sheet1")
f.SetSheetVisible("Sheet1", false)
sheets := f.GetSheetList()
```

### Streaming (for large files)
```go
sw, _ := f.NewStreamWriter("Sheet1")
sw.SetRow("A1", []interface{}{"Name", "Age"})
sw.Flush()
```

---

## Key Limitation: No .xls Support

excelize only supports .xlsx (Office Open XML). No Go library provides full .xls write support.

**Recommendation**: goxls targets .xlsx only. This is acceptable because:
- .xlsx has been the default since Excel 2007 (19 years ago)
- Most organizations have migrated
- JXLS 3.0 itself recommends .xlsx as preferred format
- .xls support can be added later if demand exists

---

## Expression Evaluation Library

### expr-lang/expr (github.com/expr-lang/expr)

**Stars**: 6,200+ | **License**: MIT | **Actively maintained**

Features:
- Safe expression evaluation (no arbitrary code execution)
- Type checking at compile time
- Fast evaluation (compiles to bytecode)
- Supports maps, structs, slices
- Supports arithmetic, comparison, logical, ternary operators
- Custom functions support

```go
import "github.com/expr-lang/expr"

env := map[string]interface{}{
    "e": Employee{Name: "Alice", Payment: 5000},
}

program, _ := expr.Compile("e.Payment > 2000", expr.Env(env))
result, _ := expr.Run(program, env)
// result = true
```

This maps well to JXLS's JEXL expressions:
- Property access: `e.Name` (Go exported fields)
- Comparisons: `e.Payment > 2000`
- Arithmetic: `e.A + e.B`
- Logical: `e.A > 0 && e.B < 100`
- Ternary: `e.Flag ? "Yes" : "No"`
- Array/slice access: `items[0]`
- Map access: `data["key"]`
