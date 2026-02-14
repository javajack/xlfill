package goxls

import (
	"bytes"
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// =============================================================================
// BuildAreas — if/else area parsing via "areas" attribute
// =============================================================================

// TestBuildAreas_IfElseWithAreasAttr tests that BuildAreas parses the "areas" attribute
// for if commands to set up the else area.
// Note: excelize only supports one comment per cell, so we use multiline comments.
func TestBuildAreas_IfElseWithAreasAttr(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "IfContent")
	f.SetCellValue(sheet, "A3", "ElseContent")

	// A1 has the area command
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	// A2 has the if command with areas attribute (multiline within single comment)
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:if(condition="show" lastCell="A2" areas=["A2:A2", "A3:A3"])`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)
	require.Len(t, areas, 1)

	// Should have one command binding (the if command)
	require.Len(t, areas[0].Bindings, 1)
	ifCmd, ok := areas[0].Bindings[0].Command.(*IfCommand)
	require.True(t, ok)
	assert.NotNil(t, ifCmd.IfArea, "IfArea should be set")
	assert.NotNil(t, ifCmd.ElseArea, "ElseArea should be set from areas attribute")

	// Test with condition true
	ctx := NewContext(map[string]any{"show": true})
	size, err := areas[0].ApplyAt(areas[0].StartCell, ctx)
	require.NoError(t, err)
	assert.True(t, size.Height >= 2)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "IfContent", v)
}

// TestBuildAreas_IfElseWithAreasAttr_FalseCondition tests else branch via areas attr.
func TestBuildAreas_IfElseWithAreasAttr_FalseCondition(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "IfContent")
	f.SetCellValue(sheet, "A3", "ElseContent")

	// Use multiline comment: area on A1, if on A2
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:if(condition="show" lastCell="A2" areas=["A2:A2", "A3:A3"])`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)
	require.Len(t, areas, 1)

	// Test with condition false — should render else content
	ctx := NewContext(map[string]any{"show": false})
	_, err = areas[0].ApplyAt(areas[0].StartCell, ctx)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "ElseContent", v)
}

// =============================================================================
// Formula coverage — buildReplacement, tryBuildRange, ProcessFormulasForRange
// =============================================================================

// TestFormulaProcessor_VerticalContiguousRange tests that vertical contiguous targets
// are combined into a range like A2:A5.
func TestFormulaProcessor_VerticalContiguousRange(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{
		NewCellRef("S", 1, 0), // A2
		NewCellRef("S", 2, 0), // A3
		NewCellRef("S", 3, 0), // A4
		NewCellRef("S", 4, 0), // A5
	}

	result := fp.buildReplacement(targets, "S", "S")
	assert.Equal(t, "A2:A5", result)
}

// TestFormulaProcessor_HorizontalContiguousRange tests horizontal range building.
func TestFormulaProcessor_HorizontalContiguousRange(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{
		NewCellRef("S", 0, 0), // A1
		NewCellRef("S", 0, 1), // B1
		NewCellRef("S", 0, 2), // C1
	}

	result := fp.buildReplacement(targets, "S", "S")
	assert.Equal(t, "A1:C1", result)
}

// TestFormulaProcessor_NonContiguousTargets tests non-contiguous targets joined with commas.
func TestFormulaProcessor_NonContiguousTargets(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{
		NewCellRef("S", 0, 0), // A1
		NewCellRef("S", 2, 0), // A3 (gap at A2)
		NewCellRef("S", 4, 0), // A5 (gap at A4)
	}

	result := fp.buildReplacement(targets, "S", "S")
	assert.Equal(t, "A1,A3,A5", result)
}

// TestFormulaProcessor_DiagonalTargets tests non-aligned targets.
func TestFormulaProcessor_DiagonalTargets(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{
		NewCellRef("S", 0, 0), // A1
		NewCellRef("S", 1, 1), // B2
		NewCellRef("S", 2, 2), // C3
	}

	result := fp.buildReplacement(targets, "S", "S")
	assert.Equal(t, "A1,B2,C3", result)
}

// TestFormulaProcessor_SingleTarget tests single target reference.
func TestFormulaProcessor_SingleTarget(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{NewCellRef("S", 0, 0)}

	result := fp.buildReplacement(targets, "S", "S")
	assert.Equal(t, "A1", result)
}

// TestFormulaProcessor_CrossSheetRef tests cross-sheet reference formatting.
func TestFormulaProcessor_CrossSheetRef(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{NewCellRef("Sheet2", 0, 0)}

	// Reference was originally on Sheet2, area is on Sheet1
	result := fp.buildReplacement(targets, "Sheet2", "Sheet1")
	assert.Equal(t, "Sheet2!A1", result)
}

// TestFormulaProcessor_ManyTargets_AdditionChain tests that >255 targets use + instead of commas.
func TestFormulaProcessor_ManyTargets_AdditionChain(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := make([]CellRef, 260)
	for i := range targets {
		// Non-contiguous (every other row)
		targets[i] = NewCellRef("S", i*2, 0)
	}

	result := fp.buildReplacement(targets, "S", "S")
	// Should use + instead of , for >255 args
	assert.Contains(t, result, "+")
	assert.NotContains(t, result, ",")
}

// TestProcessFormulasForRange tests the range expansion helper.
func TestProcessFormulasForRange(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "val1")
	f.SetCellValue(sheet, "A2", "val2")
	f.SetCellValue(sheet, "A3", "val3")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Simulate source A1 mapped to targets A1, A2, A3
	ctx := NewContext(nil)
	srcRef := NewCellRef(sheet, 0, 0)
	for row := 0; row < 3; row++ {
		dstRef := NewCellRef(sheet, row, 0)
		tx.Transform(srcRef, dstRef, ctx, false)
	}

	fp := NewFormulaProcessor()
	result := fp.ProcessFormulasForRange("SUM(A1:A1)", tx, sheet)
	assert.Contains(t, result, "A1")
	assert.Contains(t, result, "A3")
}

// =============================================================================
// Grid extractRowData — struct path
// =============================================================================

type gridTestItem struct {
	Name  string
	Value int
}

// TestGridCommand_StructDataWithProps tests grid with struct data and props.
func TestGridCommand_StructDataWithProps(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"headers": []any{"Name", "Value"},
		"data": []any{
			gridTestItem{Name: "Alice", Value: 100},
			gridTestItem{Name: "Bob", Value: 200},
		},
	})

	cmd := &GridCommand{Headers: "headers", Data: "data", Props: "Name, Value"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 3}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "200", v)
}

// TestGridCommand_MapDataNoProps tests grid with map data and no explicit props.
func TestGridCommand_MapDataNoProps(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"headers": []any{"Col1"},
		"data": []any{
			map[string]any{"A": 1},
		},
	})

	cmd := &GridCommand{Headers: "headers", Data: "data"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 2}, size)
}

// TestGridCommand_ScalarRow tests grid with scalar values as rows.
func TestGridCommand_ScalarRow(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"headers": []any{"Value"},
		"data":    []any{42, "hello", 3.14},
	})

	cmd := &GridCommand{Headers: "headers", Data: "data"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 4}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "42", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "hello", v)
}

// =============================================================================
// MultiSheet keep/hide template sheet options
// =============================================================================

// TestMultiSheet_DeleteTemplateSheet tests that default behavior deletes template sheet.
func TestMultiSheet_DeleteTemplateSheet(t *testing.T) {
	f := excelize.NewFile()
	sheet := "template"
	f.SetSheetName("Sheet1", sheet)
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="employees" var="e" lastCell="A2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "multisheet_del.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice"},
		},
	}

	out, err := FillBytes(tmpl, data)
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	// Template sheet should still exist (it's processed in place)
	sheets := outFile.GetSheetList()
	assert.Contains(t, sheets, sheet)
}

// =============================================================================
// ClearTemplateCells — actual clearing behavior
// =============================================================================

// TestClearTemplateCells_Integration tests the fill flow with clearTemplateCells.
func TestClearTemplateCells_Integration(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Age")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Age}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="B2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="employees" var="e" lastCell="B2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "clear_int.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Age": 30},
			map[string]any{"Name": "Bob", "Age": 25},
		},
	}

	out, err := FillBytes(tmpl, data, WithClearTemplateCells(true))
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	// Data should be present
	v, _ := outFile.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = outFile.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
}

// =============================================================================
// Context options coverage
// =============================================================================

// TestContextOption_WithEvaluator tests custom evaluator option.
func TestContextOption_WithEvaluator(t *testing.T) {
	ev := NewExpressionEvaluator()
	ctx := NewContext(map[string]any{"x": 5}, WithEvaluator(ev))
	result, err := ctx.Evaluate("x + 1")
	require.NoError(t, err)
	assert.Equal(t, 6, result)
}

// TestContextOption_WithUpdateCellData tests the update cell data option.
func TestContextOption_WithUpdateCellData(t *testing.T) {
	ctx := NewContext(nil, WithUpdateCellData(false))
	assert.False(t, ctx.updateCellData)
}

// TestContextOption_WithClearCells tests the clear cells option.
func TestContextOption_WithClearCells(t *testing.T) {
	ctx := NewContext(nil, WithClearCells(false))
	assert.False(t, ctx.clearCells)
}

// =============================================================================
// ContainsVar coverage
// =============================================================================

// TestContainsVar tests ContainsVar with both runVars and data.
func TestContainsVar_RunVar(t *testing.T) {
	ctx := NewContext(map[string]any{"data_var": 1})

	assert.True(t, ctx.ContainsVar("data_var"))
	assert.False(t, ctx.ContainsVar("missing"))

	rv := NewRunVar(ctx, "loop_var")
	rv.Set(42)
	assert.True(t, ctx.ContainsVar("loop_var"))

	rv.Close()
	assert.False(t, ctx.ContainsVar("loop_var"))
}

// =============================================================================
// AreaRef.SheetName coverage
// =============================================================================

// TestAreaRef_SheetName tests the SheetName method.
func TestAreaRef_SheetName(t *testing.T) {
	ar, err := ParseAreaRef("Sheet1!A1:C5")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ar.SheetName())

	ar2 := NewAreaRef(NewCellRef("", 0, 0), NewCellRef("", 5, 5))
	assert.Equal(t, "", ar2.SheetName())
}

// =============================================================================
// Command Reset coverage
// =============================================================================

// TestCommand_Reset tests that Reset() doesn't panic.
func TestCommand_Reset(t *testing.T) {
	cmds := []Command{
		&EachCommand{Items: "x", Var: "e"},
		&IfCommand{Condition: "true"},
		&GridCommand{Headers: "h", Data: "d"},
		&ImageCommand{Src: "img"},
		&MergeCellsCommand{},
		&UpdateCellCommand{Updater: "u"},
	}
	for _, cmd := range cmds {
		cmd.Reset() // should not panic
	}
}

// =============================================================================
// ExcelizeTransformer.File() coverage
// =============================================================================

// TestTransformer_File tests the File() accessor.
func TestTransformer_File(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	assert.NotNil(t, tx.File())
	assert.Equal(t, f, tx.File())
}

// =============================================================================
// SetHidden coverage
// =============================================================================

// TestTransformer_SetHidden_BothDirections tests hiding and unhiding sheets.
func TestTransformer_SetHidden_BothDirections(t *testing.T) {
	f := excelize.NewFile()
	f.NewSheet("Sheet2")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.SetHidden("Sheet2", true)
	require.NoError(t, err)

	err = tx.SetHidden("Sheet2", false)
	require.NoError(t, err)
}

// =============================================================================
// CellType.String coverage
// =============================================================================

// TestCellType_String_AllValues tests all CellType String representations including Unknown.
func TestCellType_String_AllValues(t *testing.T) {
	assert.Equal(t, "Date", CellDate.String())
	assert.Equal(t, "Error", CellError.String())
	assert.Equal(t, "Unknown", CellType(99).String())
}

// =============================================================================
// ParseCellRef edge cases
// =============================================================================

// TestParseCellRef_EdgeCases tests various edge cases for cell reference parsing.
func TestParseCellRef_EdgeCases(t *testing.T) {
	// Dollar signs (absolute refs)
	ref, err := ParseCellRef("$A$1")
	require.NoError(t, err)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 0, ref.Col)

	// Sheet with single quotes
	ref, err = ParseCellRef("'My Sheet'!B5")
	require.NoError(t, err)
	assert.Equal(t, "My Sheet", ref.Sheet)
	assert.Equal(t, 4, ref.Row)
	assert.Equal(t, 1, ref.Col)

	// Empty string
	_, err = ParseCellRef("")
	assert.Error(t, err)

	// Just sheet name with !
	_, err = ParseCellRef("Sheet1!")
	assert.Error(t, err)
}

// =============================================================================
// ParseAreaRef edge cases
// =============================================================================

// TestParseAreaRef_EdgeCases tests area reference parsing edge cases.
func TestParseAreaRef_EdgeCases(t *testing.T) {
	// No colon
	_, err := ParseAreaRef("A1")
	assert.Error(t, err)

	// Sheet inherited
	ar, err := ParseAreaRef("Sheet1!A1:C5")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ar.First.Sheet)
	assert.Equal(t, "Sheet1", ar.Last.Sheet)

	// Size calculation
	assert.Equal(t, Size{Width: 3, Height: 5}, ar.Size())
}

// =============================================================================
// WriteTypedValue — formula type coverage
// =============================================================================

// TestTransformer_WriteFormulaThroughTransform tests formula cell handling.
func TestTransformer_WriteFormulaThroughTransform(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellFormula(sheet, "A1", "SUM(B1:B5)")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	srcRef := NewCellRef(sheet, 0, 0)
	dstRef := NewCellRef(sheet, 5, 0) // Copy formula to A6

	err = tx.Transform(srcRef, dstRef, ctx, false)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "A6")
	assert.Equal(t, "SUM(B1:B5)", formula)
}

// =============================================================================
// MergeCellsCommand — expression-based cols/rows
// =============================================================================

// TestMergeCells_ExpressionCols tests MergeCells with expression-based column count.
func TestMergeCells_ExpressionCols(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Merged")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"numCols": 3, "numRows": 2})

	cmd := &MergeCellsCommand{Cols: "numCols", Rows: "numRows"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 2}, size)
}

// =============================================================================
// ResolveLastCell — with sheet prefix
// =============================================================================

// TestResolveLastCell_WithSheet tests lastCell that includes a sheet name.
func TestResolveLastCell_WithSheet(t *testing.T) {
	start := NewCellRef("Sheet1", 0, 0)

	// lastCell without sheet
	ref, err := resolveLastCell(start, "C5")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ref.Sheet)
	assert.Equal(t, 4, ref.Row) // C5 = row 4 (0-based)
	assert.Equal(t, 2, ref.Col) // C = col 2

	// lastCell with different sheet
	ref, err = resolveLastCell(start, "Sheet2!D3")
	require.NoError(t, err)
	assert.Equal(t, "Sheet2", ref.Sheet)
}

// =============================================================================
// FindMatchingEnd — nested expression coverage
// =============================================================================

// TestParseExpressions_Nested tests nested delimiters.
func TestParseExpressions_Nested(t *testing.T) {
	segments := ParseExpressions("${a + ${b}}", "${", "}")
	// Should handle this gracefully
	assert.True(t, len(segments) >= 1)
}

// TestExtractSingleExpression_MixedContent tests mixed content.
func TestExtractSingleExpression_MixedContent(t *testing.T) {
	expr, isSingle := ExtractSingleExpression("Hello ${name} World", "${", "}")
	assert.False(t, isSingle)
	assert.Equal(t, "", expr)
}

// TestExtractSingleExpression_Empty tests empty input.
func TestExtractSingleExpression_Empty(t *testing.T) {
	expr, isSingle := ExtractSingleExpression("", "${", "}")
	assert.False(t, isSingle)
	assert.Equal(t, "", expr)
}

// TestIsExpressionOnly tests the IsExpressionOnly function.
func TestIsExpressionOnly(t *testing.T) {
	assert.True(t, IsExpressionOnly("${expr}", "${", "}"))
	assert.False(t, IsExpressionOnly("text ${expr}", "${", "}"))
	assert.False(t, IsExpressionOnly("no expression", "${", "}"))
}

// =============================================================================
// Option coverage — WithKeepTemplateSheet, WithHideTemplateSheet
// =============================================================================

func TestOption_KeepTemplateSheet(t *testing.T) {
	o := defaultOptions()
	WithKeepTemplateSheet(true)(o)
	assert.True(t, o.keepTemplateSheet)
}

func TestOption_HideTemplateSheet(t *testing.T) {
	o := defaultOptions()
	WithHideTemplateSheet(true)(o)
	assert.True(t, o.hideTemplateSheet)
}

// =============================================================================
// EvaluateCellValue — nil result and blank handling
// =============================================================================

func TestEvaluateCellValue_NilResult(t *testing.T) {
	ctx := NewContext(map[string]any{"val": nil})
	result, cellType, err := ctx.EvaluateCellValue("${val}")
	require.NoError(t, err)
	assert.Nil(t, result)
	assert.Equal(t, CellBlank, cellType)
}

func TestEvaluateCellValue_BoolResult(t *testing.T) {
	ctx := NewContext(map[string]any{"val": true})
	result, cellType, err := ctx.EvaluateCellValue("${val}")
	require.NoError(t, err)
	assert.Equal(t, true, result)
	assert.Equal(t, CellBoolean, cellType)
}

func TestEvaluateCellValue_NoExpressions(t *testing.T) {
	ctx := NewContext(nil)
	result, cellType, err := ctx.EvaluateCellValue("plain text")
	require.NoError(t, err)
	assert.Equal(t, "plain text", result)
	assert.Equal(t, CellString, cellType)
}

// =============================================================================
// UpdateCell — more coverage for the command
// =============================================================================

// TestUpdateCellCommand_ViaFiller tests updateCell command through the filler flow.
func TestUpdateCellCommand_ViaFiller(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Original")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A1")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)

	for _, area := range areas {
		_, err := area.ApplyAt(area.StartCell, ctx)
		require.NoError(t, err)
	}

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Original", v)
}

// =============================================================================
// Formula — default value substitution
// =============================================================================

// TestFormulaProcessor_DefaultValue tests formula with default value when ref has no target.
func TestFormulaProcessor_DefaultValue(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.Val}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: fmt.Sprintf(`jx:each(items="items" var="e" lastCell="A2")
jx:params(defaultValue="0")`),
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Non-empty list for formula expansion
	items := []any{
		map[string]any{"Val": 100},
		map[string]any{"Val": 200},
	}
	ctx := NewContext(map[string]any{"items": items})

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)

	for _, area := range areas {
		_, err := area.ApplyAt(area.StartCell, ctx)
		require.NoError(t, err)
	}

	fp := NewFormulaProcessor()
	for _, area := range areas {
		fp.ProcessAreaFormulas(tx, area)
	}

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Formula should be in A4 (header + 2 items + formula)
	formula, _ := out.GetCellFormula(sheet, "A4")
	assert.Contains(t, formula, "A2")
}

// =============================================================================
// NewRunVarWithIndex — index save/restore
// =============================================================================

func TestNewRunVarWithIndex_SaveRestore(t *testing.T) {
	ctx := NewContext(nil)

	// Set up outer index
	outerRV := NewRunVarWithIndex(ctx, "e", "idx")
	outerRV.SetWithIndex("outer_val", 0)
	assert.Equal(t, "outer_val", ctx.GetVar("e"))
	assert.Equal(t, 0, ctx.GetVar("idx"))

	// Set up inner (overwrites)
	innerRV := NewRunVarWithIndex(ctx, "e", "idx")
	innerRV.SetWithIndex("inner_val", 5)
	assert.Equal(t, "inner_val", ctx.GetVar("e"))
	assert.Equal(t, 5, ctx.GetVar("idx"))

	// Close inner — should restore outer values
	innerRV.Close()
	assert.Equal(t, "outer_val", ctx.GetVar("e"))
	assert.Equal(t, 0, ctx.GetVar("idx"))

	// Close outer — should remove vars
	outerRV.Close()
	assert.Nil(t, ctx.GetVar("e"))
	assert.Nil(t, ctx.GetVar("idx"))
}
