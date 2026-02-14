package xlfill

import (
	"bytes"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestTransformer_OpenTemplate(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	sheets := tx.GetSheetNames()
	assert.Contains(t, sheets, "Sheet1")
}

func TestTransformer_GetCellData(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	// Header cell
	cd := tx.GetCellData(NewCellRef("Sheet1", 0, 0))
	require.NotNil(t, cd)
	assert.Equal(t, "Name", cd.Value)

	// Expression cell
	cd = tx.GetCellData(NewCellRef("Sheet1", 1, 0))
	require.NotNil(t, cd)
	assert.Equal(t, "${e.Name}", cd.Value)

	// Non-existent cell
	cd = tx.GetCellData(NewCellRef("Sheet1", 99, 99))
	assert.Nil(t, cd)
}

func TestTransformer_GetCommentedCells(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	commented := tx.GetCommentedCells()
	assert.GreaterOrEqual(t, len(commented), 2, "expected at least 2 commented cells")

	// Verify comments contain jx: commands
	hasArea := false
	hasEach := false
	for _, cd := range commented {
		if contains(cd.Comment, "jx:area") {
			hasArea = true
		}
		if contains(cd.Comment, "jx:each") {
			hasEach = true
		}
	}
	assert.True(t, hasArea, "should have jx:area comment")
	assert.True(t, hasEach, "should have jx:each comment")
}

func contains(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || len(s) > 0 && containsStr(s, substr))
}
func containsStr(s, sub string) bool {
	for i := 0; i <= len(s)-len(sub); i++ {
		if s[i:i+len(sub)] == sub {
			return true
		}
	}
	return false
}

func TestTransformer_Transform_StringValue(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{})
	src := NewCellRef("Sheet1", 0, 0) // "Name" header
	target := NewCellRef("Sheet1", 5, 0)

	err = tx.Transform(src, target, ctx, true)
	require.NoError(t, err)

	// Verify value was copied
	val, err := tx.file.GetCellValue("Sheet1", "A6")
	require.NoError(t, err)
	assert.Equal(t, "Name", val)
}

func TestTransformer_Transform_PreservesStyle(t *testing.T) {
	path := createStyledTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{})
	src := NewCellRef("Sheet1", 0, 0) // "Header" with red bold style
	target := NewCellRef("Sheet1", 5, 0)

	err = tx.Transform(src, target, ctx, true)
	require.NoError(t, err)

	// Verify style was copied (style ID should be non-zero)
	styleID, err := tx.file.GetCellStyle("Sheet1", "A6")
	require.NoError(t, err)
	assert.Greater(t, styleID, 0, "style should be preserved")
}

func TestTransformer_Transform_EvaluatesExpression(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	type Emp struct {
		Name   string
		Age    int
		Salary float64
	}
	ctx := NewContext(map[string]any{
		"e": Emp{Name: "Alice", Age: 30, Salary: 5000},
	})

	src := NewCellRef("Sheet1", 1, 0) // "${e.Name}"
	target := NewCellRef("Sheet1", 5, 0)

	err = tx.Transform(src, target, ctx, true)
	require.NoError(t, err)

	val, err := tx.file.GetCellValue("Sheet1", "A6")
	require.NoError(t, err)
	assert.Equal(t, "Alice", val)
}

func TestTransformer_Transform_FormulaCell(t *testing.T) {
	path := createFormulaTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{})
	// A3 has formula SUM(A2:A2)
	src := NewCellRef("Sheet1", 2, 0)
	target := NewCellRef("Sheet1", 10, 0)

	err = tx.Transform(src, target, ctx, true)
	require.NoError(t, err)

	formula, err := tx.file.GetCellFormula("Sheet1", "A11")
	require.NoError(t, err)
	assert.Equal(t, "SUM(A2:A2)", formula)
}

func TestTransformer_ClearCell(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.ClearCell(NewCellRef("Sheet1", 0, 0))
	require.NoError(t, err)

	val, err := tx.file.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "", val)
}

func TestTransformer_SetFormula(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.SetFormula(NewCellRef("Sheet1", 5, 0), "SUM(A1:A3)")
	require.NoError(t, err)

	formula, err := tx.file.GetCellFormula("Sheet1", "A6")
	require.NoError(t, err)
	assert.Equal(t, "SUM(A1:A3)", formula)
}

func TestTransformer_TrackTargetCellRef(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{})
	src := NewCellRef("Sheet1", 0, 0)
	t1 := NewCellRef("Sheet1", 5, 0)
	t2 := NewCellRef("Sheet1", 6, 0)

	tx.Transform(src, t1, ctx, false)
	tx.Transform(src, t2, ctx, false)

	targets := tx.GetTargetCellRef(src)
	assert.Len(t, targets, 2)
	assert.Equal(t, t1, targets[0])
	assert.Equal(t, t2, targets[1])
}

func TestTransformer_ResetTargetCellRefs(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{})
	src := NewCellRef("Sheet1", 0, 0)
	tx.Transform(src, NewCellRef("Sheet1", 5, 0), ctx, false)

	tx.ResetTargetCellRefs()
	targets := tx.GetTargetCellRef(src)
	assert.Empty(t, targets)
}

func TestTransformer_Write(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	var buf bytes.Buffer
	err = tx.Write(&buf)
	require.NoError(t, err)
	assert.Greater(t, buf.Len(), 0, "output should not be empty")

	// Verify output is valid xlsx
	f, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer f.Close()
	val, err := f.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Name", val)
}

func TestTransformer_DeleteSheet(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	// Add a second sheet so we can delete
	tx.file.NewSheet("Sheet2")
	err = tx.DeleteSheet("Sheet2")
	require.NoError(t, err)

	sheets := tx.GetSheetNames()
	assert.NotContains(t, sheets, "Sheet2")
}

func TestTransformer_SetHidden(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	// Need a second sheet and make it active (can't hide the active sheet)
	_, err = tx.file.NewSheet("Sheet2")
	require.NoError(t, err)
	tx.file.SetActiveSheet(1) // make Sheet2 active

	err = tx.SetHidden("Sheet1", true)
	require.NoError(t, err)

	// Write and re-read to verify
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	f2, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer f2.Close()

	visible, err := f2.GetSheetVisible("Sheet1")
	require.NoError(t, err)
	assert.False(t, visible)
}

func TestTransformer_ColumnWidth(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	// Set a known width and verify
	tx.file.SetColWidth("Sheet1", "A", "A", 20.0)
	w := tx.GetColumnWidth("Sheet1", 0)
	assert.InDelta(t, 20.0, w, 0.5)
}

func TestTransformer_RowHeight(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	tx.file.SetRowHeight("Sheet1", 1, 30.0)
	h := tx.GetRowHeight("Sheet1", 0) // 0-based
	assert.InDelta(t, 30.0, h, 0.5)
}

func TestTransformer_MergeCells(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.MergeCells("Sheet1", "A5", "C5")
	require.NoError(t, err)

	merged, err := tx.file.GetMergeCells("Sheet1")
	require.NoError(t, err)

	found := false
	for _, m := range merged {
		if m.GetStartAxis() == "A5" && m.GetEndAxis() == "C5" {
			found = true
		}
	}
	assert.True(t, found, "merged cell range A5:C5 should exist")
}

func TestTransformer_GetFormulaCells(t *testing.T) {
	path := createFormulaTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	formulas := tx.GetFormulaCells()
	assert.GreaterOrEqual(t, len(formulas), 1, "should find at least one formula cell")

	hasSum := false
	for _, cd := range formulas {
		if cd.Formula == "SUM(A2:A2)" {
			hasSum = true
		}
	}
	assert.True(t, hasSum, "should find SUM formula")
}

func TestTransformer_NonExistentFile(t *testing.T) {
	_, err := OpenTemplate("/nonexistent/path.xlsx")
	assert.Error(t, err)
}

func TestTransformer_SetCellValue(t *testing.T) {
	path := createBasicTemplate(t)
	defer os.Remove(path)

	tx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.SetCellValue(NewCellRef("Sheet1", 0, 0), "NewValue")
	require.NoError(t, err)

	val, err := tx.file.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "NewValue", val)
}
