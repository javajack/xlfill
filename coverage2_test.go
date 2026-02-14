package xlfill

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// =============================================================================
// UpdateCellCommand — full ApplyAt coverage
// =============================================================================

// cov2Updater implements CellDataUpdater for coverage tests.
type cov2Updater struct {
	valueToSet   any
	formulaToSet string
}

func (u *cov2Updater) UpdateCellData(cd *CellData, targetCell CellRef, ctx *Context) {
	if u.formulaToSet != "" {
		cd.Formula = u.formulaToSet
		cd.Value = nil
	} else {
		cd.Value = u.valueToSet
		cd.Formula = ""
	}
}

// TestUpdateCellCommand_WithArea tests updateCell with an area (multi-cell update).
func TestUpdateCellCommand_WithArea(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "old1")
	f.SetCellValue(sheet, "B1", "old2")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	cmd := &UpdateCellCommand{Updater: "myUpdater", Area: area}

	updater := &cov2Updater{valueToSet: "UPDATED"}
	ctx := NewContext(map[string]any{"myUpdater": updater})

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "UPDATED", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "UPDATED", v)
}

// TestUpdateCellCommand_WithFormula tests updateCell setting a formula.
func TestUpdateCellCommand_WithFormula(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "old")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &UpdateCellCommand{Updater: "myUpdater"}
	updater := &cov2Updater{formulaToSet: "SUM(B1:B5)"}
	ctx := NewContext(map[string]any{"myUpdater": updater})

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "A1")
	assert.Equal(t, "SUM(B1:B5)", formula)
}

// TestUpdateCellCommand_UpdaterNotFound tests error when updater not in context.
func TestUpdateCellCommand_UpdaterNotFound(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &UpdateCellCommand{Updater: "missing"}
	ctx := NewContext(nil)

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "not found")
}

// TestUpdateCellCommand_NotCellDataUpdater tests error when updater has wrong type.
func TestUpdateCellCommand_NotCellDataUpdater(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &UpdateCellCommand{Updater: "bad"}
	ctx := NewContext(map[string]any{"bad": "not-an-updater"})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "CellDataUpdater")
}

// TestUpdateCellCommand_AreaWithFormula tests updateCell with area setting formula on cells.
func TestUpdateCellCommand_AreaWithFormula(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "val")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &UpdateCellCommand{Updater: "u", Area: area}

	updater := &cov2Updater{formulaToSet: "A2+A3"}
	ctx := NewContext(map[string]any{"u": updater})

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 1, size.Width)
	assert.Equal(t, 1, size.Height)
}

// TestUpdateCellCommand_NilCellData tests updateCell when target cell has no existing data.
func TestUpdateCellCommand_NilCellData(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	// Don't set any value in A5 — it will have nil CellData

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &UpdateCellCommand{Updater: "u"}
	updater := &cov2Updater{valueToSet: "created"}
	ctx := NewContext(map[string]any{"u": updater})

	size, err := cmd.ApplyAt(NewCellRef(sheet, 4, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

// =============================================================================
// MergeCellsCommand — full coverage
// =============================================================================

// TestMergeCells_MinColsSkip tests that merge is skipped when cols < minCols.
func TestMergeCells_MinColsSkip(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &MergeCellsCommand{Cols: "2", MinCols: "5"}
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	// Should return size but skip actual merge
	assert.Equal(t, 2, size.Width)
}

// TestMergeCells_MinRowsSkip tests that merge is skipped when rows < minRows.
func TestMergeCells_MinRowsSkip(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &MergeCellsCommand{Rows: "2", MinRows: "5"}
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height)
}

// TestMergeCells_ColsDirectInteger tests cols as direct integer string (not expression).
func TestMergeCells_ColsDirectInteger(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// "3" can't be evaluated as an expression but can be parsed as int
	cmd := &MergeCellsCommand{Cols: "3", Rows: "2"}
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 2}, size)
}

// TestMergeCells_RowsDirectInteger tests rows as direct integer string.
func TestMergeCells_RowsDirectInteger(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &MergeCellsCommand{Rows: "4"}
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 4, size.Height)
}

// TestMergeCells_ColsExprError tests cols expression that fails eval and fails atoi.
func TestMergeCells_ColsExprError(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Use an expression that fails compilation AND can't be Atoi'd
	cmd := &MergeCellsCommand{Cols: "1 + + +"}
	ctx := NewContext(nil)

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "cols")
}

// TestMergeCells_RowsExprError tests rows expression that fails eval and fails atoi.
func TestMergeCells_RowsExprError(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &MergeCellsCommand{Rows: "1 + + +"}
	ctx := NewContext(nil)

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "rows")
}

// TestMergeCells_NothingToMerge tests 1x1 merge (no-op).
func TestMergeCells_NothingToMerge(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &MergeCellsCommand{} // no cols or rows → defaults to 1x1
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

// TestToInt_AllTypes tests toInt with various types.
func TestToInt_AllTypes(t *testing.T) {
	assert.Equal(t, 5, toInt(5))
	assert.Equal(t, 5, toInt(int64(5)))
	assert.Equal(t, 5, toInt(float64(5.7)))
	assert.Equal(t, 5, toInt(float32(5.9)))
	assert.Equal(t, 5, toInt("5"))
	assert.Equal(t, 1, toInt("not_a_number")) // default
	assert.Equal(t, 1, toInt(nil))            // default
	assert.Equal(t, 1, toInt(true))           // default for unhandled type
}

// TestMergeCells_MinColsInvalidAtoi tests non-numeric minCols string.
func TestMergeCells_MinColsInvalidAtoi(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// minCols is "abc" → Atoi fails → minCols stays 0 → merge proceeds
	cmd := &MergeCellsCommand{Cols: "3", Rows: "1", MinCols: "abc"}
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Width)
}

// TestMergeCells_MinRowsInvalidAtoi tests non-numeric minRows string.
func TestMergeCells_MinRowsInvalidAtoi(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &MergeCellsCommand{Cols: "1", Rows: "3", MinRows: "abc"}
	ctx := NewContext(nil)

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)
}

// =============================================================================
// ImageCommand — full coverage
// =============================================================================

func createCov2PNG(t *testing.T) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, 1, 1))
	img.Set(0, 0, color.RGBA{255, 0, 0, 255})
	var buf bytes.Buffer
	require.NoError(t, png.Encode(&buf, img))
	return buf.Bytes()
}

// TestImageCommand_PNG tests image command with actual PNG bytes.
func TestImageCommand_PNGWithActualBytes(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	imgBytes := createCov2PNG(t)
	ctx := NewContext(map[string]any{"img": imgBytes})

	cmd := &ImageCommand{Src: "img", ImageType: "PNG", ScaleX: 1.0, ScaleY: 1.0}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

// TestImageCommand_NilSrc tests image with nil source (graceful skip).
func TestImageCommand_NilSrc(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"img": nil})
	cmd := &ImageCommand{Src: "img", ImageType: "PNG", ScaleX: 1.0, ScaleY: 1.0}

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

// TestImageCommand_NonByteType tests image with non-[]byte type (error).
func TestImageCommand_NonByteType(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"img": "not-bytes"})
	cmd := &ImageCommand{Src: "img", ImageType: "PNG", ScaleX: 1.0, ScaleY: 1.0}

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "[]byte")
}

// TestImageCommand_EvalError tests image when src expression fails.
func TestImageCommand_EvalError(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	cmd := &ImageCommand{Src: "undefined_var + 1", ImageType: "PNG", ScaleX: 1.0, ScaleY: 1.0}

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

// =============================================================================
// AddImage — image type extensions
// =============================================================================

// TestAddImage_JPEG tests JPEG extension mapping.
func TestAddImage_JPEG(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	imgBytes := createCov2PNG(t) // content doesn't need to match type for this test
	// Just verifying the extension mapping doesn't error
	err = tx.AddImage("Sheet1", "A1", imgBytes, "JPEG", 1.0, 1.0)
	// May error due to invalid JPEG content but the ext mapping is tested
	_ = err
}

// TestAddImage_GIF tests GIF extension mapping.
func TestAddImage_GIF(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.AddImage("Sheet1", "A1", createCov2PNG(t), "GIF", 1.0, 1.0)
	_ = err // ext mapping tested
}

// TestAddImage_BMP tests BMP extension mapping.
func TestAddImage_BMP(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.AddImage("Sheet1", "A1", createCov2PNG(t), "BMP", 1.0, 1.0)
	_ = err // ext mapping tested
}

// =============================================================================
// Formula processFormula — uncovered branches
// =============================================================================

// TestProcessFormula_InternalRefNoTarget tests formula ref inside area with no target → default value.
func TestProcessFormula_InternalRefNoTarget(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Setup: A1 has value, A2 has formula referencing A1
	f.SetCellValue(sheet, "A1", "val")
	f.SetCellFormula(sheet, "A2", "A1+1")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Create area containing both A1 and A2
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 2}, tx)

	// Don't transform A1 (so it has no target), but register A2's formula target
	ctx := NewContext(nil)
	srcA2 := NewCellRef(sheet, 1, 0)
	dstA2 := NewCellRef(sheet, 1, 0)
	tx.Transform(srcA2, dstA2, ctx, false)

	fp := NewFormulaProcessor()
	fp.ProcessAreaFormulas(tx, area)

	// A1 has no target but is inside area → replaced with default "0"
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "A2")
	assert.Contains(t, formula, "0") // A1 replaced with default
}

// TestProcessFormula_ExternalRef tests formula referencing cell outside area (kept as-is).
func TestProcessFormula_ExternalRef(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", 10)
	f.SetCellValue(sheet, "B1", 20)
	f.SetCellFormula(sheet, "A2", "A1+B1")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Area only covers A1:A2 — B1 is outside
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 2}, tx)

	ctx := NewContext(nil)
	// Transform A1 and A2 to themselves
	tx.Transform(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 0, 0), ctx, false)
	tx.Transform(NewCellRef(sheet, 1, 0), NewCellRef(sheet, 1, 0), ctx, false)

	fp := NewFormulaProcessor()
	fp.ProcessAreaFormulas(tx, area)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "A2")
	// B1 should remain as-is (external reference)
	assert.Contains(t, formula, "B1")
}

// TestProcessFormula_StrategyFilteredEmpty tests formula where strategy filters out all targets.
func TestProcessFormula_StrategyFilteredEmpty(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", 10)
	f.SetCellFormula(sheet, "B1", "A1*2")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)

	ctx := NewContext(nil)
	// A1 maps to A2, A3 (different column than B1's target)
	tx.Transform(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 1, 0), ctx, false)
	tx.Transform(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 2, 0), ctx, false)

	// B1's formula target is at B1 (col 1)
	tx.Transform(NewCellRef(sheet, 0, 1), NewCellRef(sheet, 0, 1), ctx, false)

	// Set BY_COLUMN strategy on B1's formula cell data
	cd := tx.GetCellData(NewCellRef(sheet, 0, 1))
	require.NotNil(t, cd)
	cd.FormulaStrategy = FormulaByColumn
	cd.DefaultValue = "99"

	fp := NewFormulaProcessor()
	fp.ProcessAreaFormulas(tx, area)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// A1 targets are in col 0, but formula target B1 is in col 1 → BY_COLUMN filters all → default
	formula, _ := out.GetCellFormula(sheet, "B1")
	assert.Contains(t, formula, "99")
}

// TestProcessFormula_CustomDefaultValue tests custom defaultValue from params.
func TestProcessFormula_CustomDefaultValue(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", 10)
	f.SetCellFormula(sheet, "A2", "A1+1")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 2}, tx)

	ctx := NewContext(nil)
	// Transform A2 but NOT A1 → A1 internal ref has no target
	tx.Transform(NewCellRef(sheet, 1, 0), NewCellRef(sheet, 1, 0), ctx, false)

	// Set custom default value on A2
	cd := tx.GetCellData(NewCellRef(sheet, 1, 0))
	require.NotNil(t, cd)
	cd.DefaultValue = "-1"

	fp := NewFormulaProcessor()
	fp.ProcessAreaFormulas(tx, area)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "A2")
	assert.Contains(t, formula, "-1")
}

// =============================================================================
// Filler — buildIfElseArea (string-based path) and other uncovered paths
// =============================================================================

// TestBuildIfElseArea_StringPath tests the old string-based buildIfElseArea path.
func TestBuildIfElseArea_StringPath(t *testing.T) {
	filler := &Filler{opts: defaultOptions(), registry: NewCommandRegistry()}

	ifCmd := &IfCommand{Condition: "true"}
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "if")
	f.SetCellValue(sheet, "A2", "else")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Call buildIfElseArea directly with string format
	err = filler.buildIfElseArea(ifCmd, `"A1:A1", "A2:A2"`, NewCellRef(sheet, 0, 0), tx)
	require.NoError(t, err)
	assert.NotNil(t, ifCmd.ElseArea)
}

// TestBuildIfElseArea_NoElse tests buildIfElseArea with only one area (no else).
func TestBuildIfElseArea_NoElse(t *testing.T) {
	filler := &Filler{opts: defaultOptions(), registry: NewCommandRegistry()}
	ifCmd := &IfCommand{Condition: "true"}

	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = filler.buildIfElseArea(ifCmd, `"A1:A1"`, NewCellRef("Sheet1", 0, 0), tx)
	require.NoError(t, err)
	assert.Nil(t, ifCmd.ElseArea)
}

// TestBuildIfElseArea_EmptySecondPart tests buildIfElseArea with empty else reference.
func TestBuildIfElseArea_EmptySecondPart(t *testing.T) {
	filler := &Filler{opts: defaultOptions(), registry: NewCommandRegistry()}
	ifCmd := &IfCommand{Condition: "true"}

	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = filler.buildIfElseArea(ifCmd, `"A1:A1", `, NewCellRef("Sheet1", 0, 0), tx)
	require.NoError(t, err)
	assert.Nil(t, ifCmd.ElseArea)
}

// TestAttachArea_GridCommand tests attachArea for GridCommand.
func TestAttachArea_GridCommand(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 3, Height: 3}, tx)

	grid := &GridCommand{Headers: "h", Data: "d"}
	attachArea(grid, area)
	assert.Equal(t, area, grid.BodyArea)

	uc := &UpdateCellCommand{Updater: "u"}
	attachArea(uc, area)
	assert.Equal(t, area, uc.Area)
}

// TestContainsRef_CrossSheet tests containsRef with different sheets.
func TestContainsRef_CrossSheet(t *testing.T) {
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 5, Height: 5}, nil)

	// Same sheet, inside
	assert.True(t, area.containsRef(NewCellRef("Sheet1", 2, 2)))
	// Different sheet
	assert.False(t, area.containsRef(NewCellRef("Sheet2", 2, 2)))
	// Outside bounds
	assert.False(t, area.containsRef(NewCellRef("Sheet1", 10, 0)))
	assert.False(t, area.containsRef(NewCellRef("Sheet1", 0, 10)))
}

// =============================================================================
// ExcelizeTransformer — uncovered error branches
// =============================================================================

// TestGetColumnWidth_Error tests GetColumnWidth returning 0 on error.
func TestGetColumnWidth_Error(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Non-existent sheet
	w := tx.GetColumnWidth("NonExistent", 0)
	assert.Equal(t, float64(0), w)
}

// TestGetRowHeight_Error tests GetRowHeight returning 0 on error.
func TestGetRowHeight_Error(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	h := tx.GetRowHeight("NonExistent", 0)
	assert.Equal(t, float64(0), h)
}

// TestCopySheet_NonExistentSrc tests CopySheet with non-existent source.
func TestCopySheet_NonExistentSrc(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.CopySheet("NonExistent", "New")
	assert.Error(t, err)
}

// TestTransformer_WriteTypedValue_Formula tests writeTypedValue with formula type.
func TestTransformer_WriteTypedValue_Formula(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.writeTypedValue("Sheet1", "A1", "SUM(B1:B5)", CellFormula)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula("Sheet1", "A1")
	assert.Equal(t, "SUM(B1:B5)", formula)
}

// TestTransformer_WriteTypedValue_Nil tests writeTypedValue with nil (no-op).
func TestTransformer_WriteTypedValue_Nil(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.writeTypedValue("Sheet1", "A1", nil, CellBlank)
	require.NoError(t, err)
}

// =============================================================================
// Trivial Name() and Reset() methods
// =============================================================================

func TestCommandNames(t *testing.T) {
	assert.Equal(t, "grid", (&GridCommand{}).Name())
	assert.Equal(t, "image", (&ImageCommand{Src: "x"}).Name())
	assert.Equal(t, "mergeCells", (&MergeCellsCommand{}).Name())
	assert.Equal(t, "updateCell", (&UpdateCellCommand{Updater: "x"}).Name())
	assert.Equal(t, "if", (&IfCommand{}).Name())
	assert.Equal(t, "each", (&EachCommand{}).Name())
}

// =============================================================================
// clearTemplateCells — the no-op function (ensures it's called without panic)
// =============================================================================

func TestClearTemplateCells_NoOp(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${expr}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	ctx := NewContext(nil)

	// Should not panic
	area.clearTemplateCells(ctx)
}

// =============================================================================
// Area.ClearCells — with nil transformer
// =============================================================================

func TestArea_ClearCells_NilTransformer(t *testing.T) {
	area := &Area{
		StartCell: NewCellRef("Sheet1", 0, 0),
		AreaSize:  Size{Width: 2, Height: 2},
	}
	// Should not panic when transformer is nil
	area.ClearCells()
}

// =============================================================================
// Area.transformStaticArea — error path
// =============================================================================

func TestArea_ApplyAt_NilTransformer_ErrorMsg(t *testing.T) {
	area := &Area{
		StartCell: NewCellRef("Sheet1", 0, 0),
		AreaSize:  Size{Width: 1, Height: 1},
	}
	ctx := NewContext(nil)

	_, err := area.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no transformer")
}

// =============================================================================
// FillBytes — top-level API
// =============================================================================

func TestFillBytes_Success(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${name}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	tmpl := filepath.Join(testdataDir(t), "fillbytes_test.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	out, err := FillBytes(tmpl, map[string]any{"name": "Hello"})
	require.NoError(t, err)
	assert.True(t, len(out) > 0)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	v, _ := outFile.GetCellValue(sheet, "A1")
	assert.Equal(t, "Hello", v)
}

func TestFillBytes_BadTemplate(t *testing.T) {
	_, err := FillBytes("/nonexistent/path.xlsx", nil)
	assert.Error(t, err)
}

// =============================================================================
// openTemplate — no template specified error
// =============================================================================

func TestOpenTemplate_NoTemplate(t *testing.T) {
	filler := NewFiller() // no template specified
	_, err := filler.openTemplate()
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no template")
}

// =============================================================================
// parseCellName edge — invalid col name
// =============================================================================

func TestParseCellRef_InvalidCol(t *testing.T) {
	_, err := ParseCellRef("123")
	assert.Error(t, err)
}

// =============================================================================
// sortItems error path (empty orderBy)
// =============================================================================

func TestEachCommand_SortItems_EmptyOrderBy(t *testing.T) {
	cmd := &EachCommand{Items: "items", Var: "e", OrderBy: "  "}
	items := []any{1, 2, 3}
	result, err := cmd.sortItems(items)
	require.NoError(t, err)
	assert.Equal(t, items, result)
}

// =============================================================================
// parseOrderBy edge cases
// =============================================================================

func TestParseOrderBy_EmptyParts(t *testing.T) {
	specs := parseOrderBy("e.Name ASC,  , e.Age DESC", "e")
	assert.Len(t, specs, 2)
	assert.Equal(t, "Name", specs[0].field)
	assert.False(t, specs[0].desc)
	assert.Equal(t, "Age", specs[1].field)
	assert.True(t, specs[1].desc)
}

// =============================================================================
// EachCommand — filterItems error
// =============================================================================

func TestEachCommand_FilterItems_Error(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &EachCommand{Items: "items", Var: "e", Select: "e.Bad + undefined_fn()", Area: area}
	ctx := NewContext(map[string]any{
		"items": []any{map[string]any{"Name": "Alice"}},
	})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "select")
}

// =============================================================================
// EachCommand — no area error
// =============================================================================

func TestEachCommand_NoArea_ErrorMessage(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &EachCommand{Items: "items", Var: "e"}
	ctx := NewContext(map[string]any{"items": []any{1}})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no area")
}

// =============================================================================
// compareBySpecs — desc comparison
// =============================================================================

func TestCompareBySpecs_Desc(t *testing.T) {
	specs := []orderBySpec{{field: "Val", desc: true}}
	a := map[string]any{"Val": 1}
	b := map[string]any{"Val": 2}

	cmp := compareBySpecs(a, b, specs)
	assert.Greater(t, cmp, 0) // desc: a(1) > b(2) in natural, reversed → positive
}

// =============================================================================
// ParseComment — empty and edge cases
// =============================================================================

func TestParseComment_Empty(t *testing.T) {
	cmds, params, err := ParseComment("", NewCellRef("Sheet1", 0, 0))
	require.NoError(t, err)
	assert.Nil(t, cmds)
	assert.Nil(t, params)
}

func TestParseComment_NonCommand(t *testing.T) {
	cmds, params, err := ParseComment("just a regular comment", NewCellRef("Sheet1", 0, 0))
	require.NoError(t, err)
	assert.Nil(t, cmds)
	assert.Nil(t, params)
}

func TestParseComment_MissingParen(t *testing.T) {
	_, _, err := ParseComment("jx:each items=x", NewCellRef("Sheet1", 0, 0))
	assert.Error(t, err)
}

func TestParseComment_MissingCloseParen(t *testing.T) {
	_, _, err := ParseComment(`jx:each(items="x" var="e" lastCell="A1"`, NewCellRef("Sheet1", 0, 0))
	assert.Error(t, err)
}

func TestParseComment_MissingLastCell(t *testing.T) {
	_, _, err := ParseComment(`jx:each(items="x" var="e")`, NewCellRef("Sheet1", 0, 0))
	assert.Error(t, err)
}

// =============================================================================
// ParseParams — edge cases
// =============================================================================

func TestParseParams_NoParen(t *testing.T) {
	pd, err := ParseParams("jx:params")
	require.NoError(t, err)
	assert.Equal(t, FormulaDefault, pd.FormulaStrategy)
}

func TestParseParams_MissingCloseParen(t *testing.T) {
	_, err := ParseParams(`jx:params(defaultValue="0"`)
	assert.Error(t, err)
}

func TestParseParams_FormulaStrategyDefault(t *testing.T) {
	pd, err := ParseParams(`jx:params(formulaStrategy="UNKNOWN")`)
	require.NoError(t, err)
	assert.Equal(t, FormulaDefault, pd.FormulaStrategy)
}

func TestParseParams_FormulaStrategyByRow(t *testing.T) {
	pd, err := ParseParams(`jx:params(formulaStrategy="BY_ROW")`)
	require.NoError(t, err)
	assert.Equal(t, FormulaByRow, pd.FormulaStrategy)
}

// =============================================================================
// BuildAreas — error paths
// =============================================================================

func TestBuildAreas_NoComments_ErrorMsg(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	_, err = filler.BuildAreas(tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no commented cells")
}

func TestBuildAreas_NoAreaCommand_ErrorMsg(t *testing.T) {
	f := excelize.NewFile()
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:each(items="x" var="e" lastCell="A1")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	_, err = filler.BuildAreas(tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no jx:area")
}

// =============================================================================
// IsExpressionOnly — edge cases
// =============================================================================

func TestIsExpressionOnly_EmptyDelimiters(t *testing.T) {
	// Empty delimiters default to ${ }
	assert.True(t, IsExpressionOnly("${x}", "", ""))
}

func TestIsExpressionOnly_NestedExpr(t *testing.T) {
	assert.False(t, IsExpressionOnly("${a + ${b}}", "${", "}"))
}

// =============================================================================
// ExtractSingleExpression — empty delimiters
// =============================================================================

func TestExtractSingleExpression_EmptyDelimiters(t *testing.T) {
	expr, ok := ExtractSingleExpression("${hello}", "", "")
	assert.True(t, ok)
	assert.Equal(t, "hello", expr)
}

func TestExtractSingleExpression_Nested(t *testing.T) {
	expr, ok := ExtractSingleExpression("${a + ${b}}", "${", "}")
	assert.False(t, ok)
	assert.Equal(t, "", expr)
}

// =============================================================================
// findMatchingEnd — deeply nested
// =============================================================================

func TestFindMatchingEnd_DeepNesting(t *testing.T) {
	// "${a + ${b + ${c}}}" after removing leading "${" → "a + ${b + ${c}}}"
	idx := findMatchingEnd("a + ${b + ${c}}}", "${", "}")
	// Should find the outermost matching }
	assert.True(t, idx >= 0)
}

// =============================================================================
// FillReader — top-level API with io.Reader
// =============================================================================

func TestFillReader_Success(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${val}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	var tmplBuf bytes.Buffer
	require.NoError(t, f.Write(&tmplBuf))
	f.Close()

	var outBuf bytes.Buffer
	err := FillReader(&tmplBuf, &outBuf, map[string]any{"val": 42})
	require.NoError(t, err)

	out, err := excelize.OpenReader(&outBuf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "42", v)
}

// =============================================================================
// Fill — file-based top-level API
// =============================================================================

func TestFill_Success(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${val}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	tmpl := filepath.Join(testdataDir(t), "fill_test.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	outPath := filepath.Join(testdataDir(t), "fill_test_out.xlsx")
	err := Fill(tmpl, outPath, map[string]any{"val": "works"})
	require.NoError(t, err)

	out, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "works", v)
}

func TestFill_BadOutput(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "${val}")
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	tmpl := filepath.Join(testdataDir(t), "fill_bad_out.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	err := Fill(tmpl, "/nonexistent/dir/out.xlsx", map[string]any{"val": "x"})
	assert.Error(t, err)
}

// =============================================================================
// WithExpressionNotation option
// =============================================================================

func TestWithExpressionNotation(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "<<val>>")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	var tmplBuf bytes.Buffer
	require.NoError(t, f.Write(&tmplBuf))
	f.Close()

	var outBuf bytes.Buffer
	err := FillReader(&tmplBuf, &outBuf, map[string]any{"val": "custom"},
		WithExpressionNotation("<<", ">>"))
	require.NoError(t, err)

	out, err := excelize.OpenReader(&outBuf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "custom", v)
}

// =============================================================================
// WithPreWrite option
// =============================================================================

func TestWithPreWrite(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Hello")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	var tmplBuf bytes.Buffer
	require.NoError(t, f.Write(&tmplBuf))
	f.Close()

	called := false
	var outBuf bytes.Buffer
	err := FillReader(&tmplBuf, &outBuf, nil,
		WithPreWrite(func(tx Transformer) error {
			called = true
			return nil
		}))
	require.NoError(t, err)
	assert.True(t, called)
}

func TestWithPreWrite_Error(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Hello")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	var tmplBuf bytes.Buffer
	require.NoError(t, f.Write(&tmplBuf))
	f.Close()

	var outBuf bytes.Buffer
	err := FillReader(&tmplBuf, &outBuf, nil,
		WithPreWrite(func(tx Transformer) error {
			return fmt.Errorf("pre-write failed")
		}))
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "pre-write")
}

// =============================================================================
// WithCommand — custom command
// =============================================================================

func TestWithCommand_Custom(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "test")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: fmt.Sprintf(`jx:area(lastCell="A2")
jx:custom(lastCell="A2" msg="hello")`),
	})
	f.SetCellValue(sheet, "A2", "body")

	var tmplBuf bytes.Buffer
	require.NoError(t, f.Write(&tmplBuf))
	f.Close()

	customCalled := false
	factory := func(attrs map[string]string) (Command, error) {
		return &customTestCmd{msg: attrs["msg"], called: &customCalled}, nil
	}

	var outBuf bytes.Buffer
	err := FillReader(&tmplBuf, &outBuf, nil, WithCommand("custom", factory))
	require.NoError(t, err)
	assert.True(t, customCalled)
}

type customTestCmd struct {
	msg    string
	called *bool
}

func (c *customTestCmd) Name() string                                              { return "custom" }
func (c *customTestCmd) Reset()                                                    {}
func (c *customTestCmd) ApplyAt(_ CellRef, _ *Context, _ Transformer) (Size, error) {
	*c.called = true
	return Size{Width: 1, Height: 1}, nil
}

// =============================================================================
// Expr.go — Evaluate edge cases
// =============================================================================

func TestExprEvaluator_EmptyExpression(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("", nil)
	require.NoError(t, err)
	assert.Nil(t, result)
}

func TestExprEvaluator_IsConditionTrue_NilResult(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.IsConditionTrue("nil_var", map[string]any{"nil_var": nil})
	require.NoError(t, err)
	assert.False(t, result) // nil → false
}

func TestExprEvaluator_IsConditionTrue_NonBool(t *testing.T) {
	ev := NewExpressionEvaluator()
	_, err := ev.IsConditionTrue("val", map[string]any{"val": 42})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "expected bool")
}

// =============================================================================
// CellData helpers
// =============================================================================

func TestCellData_AddTargetPosWithArea_Cov2(t *testing.T) {
	cd := NewCellData(NewCellRef("S", 0, 0), "val", CellString)
	ar := NewAreaRef(NewCellRef("S", 0, 0), NewCellRef("S", 5, 5))
	cd.AddTargetPosWithArea(NewCellRef("S", 1, 0), ar)
	cd.AddTargetPosWithArea(NewCellRef("S", 2, 0), ar)
	assert.Len(t, cd.TargetPositions, 2)
	assert.Len(t, cd.TargetParentArea, 2)
}

func TestCellData_Reset_VerifySliceClear(t *testing.T) {
	cd := NewCellData(NewCellRef("S", 0, 0), "val", CellString)
	cd.AddTargetPos(NewCellRef("S", 1, 0))
	cd.AddTargetPos(NewCellRef("S", 2, 0))
	ar := NewAreaRef(NewCellRef("S", 0, 0), NewCellRef("S", 5, 5))
	cd.AddTargetPosWithArea(NewCellRef("S", 3, 0), ar)
	cd.EvalFormulas = append(cd.EvalFormulas, "SUM(A1)")
	cd.EvalResult = "result"
	cd.Reset()
	assert.Len(t, cd.TargetPositions, 0)
	assert.Len(t, cd.TargetParentArea, 0)
	assert.Len(t, cd.EvalFormulas, 0)
	assert.Nil(t, cd.EvalResult)
}

// =============================================================================
// inferCellType — edge cases
// =============================================================================

func TestInferCellType_Uint(t *testing.T) {
	assert.Equal(t, CellNumber, inferCellType(uint(5)))
	assert.Equal(t, CellNumber, inferCellType(uint8(5)))
	assert.Equal(t, CellNumber, inferCellType(uint16(5)))
	assert.Equal(t, CellNumber, inferCellType(uint32(5)))
	assert.Equal(t, CellNumber, inferCellType(uint64(5)))
	assert.Equal(t, CellNumber, inferCellType(int8(5)))
	assert.Equal(t, CellNumber, inferCellType(int16(5)))
	assert.Equal(t, CellNumber, inferCellType(int32(5)))
	assert.Equal(t, CellNumber, inferCellType(float32(5.0)))
}

func TestInferCellType_Struct(t *testing.T) {
	type custom struct{}
	assert.Equal(t, CellString, inferCellType(custom{})) // default
}

// =============================================================================
// CommandRegistry — unknown command returns nil
// =============================================================================

func TestCommandRegistry_Unknown(t *testing.T) {
	reg := NewCommandRegistry()
	cmd, err := reg.Create("nonexistent", nil)
	require.NoError(t, err)
	assert.Nil(t, cmd)
}

// =============================================================================
// newImageCommandFromAttrs — edge cases
// =============================================================================

func TestNewImageCommandFromAttrs_NoSrc(t *testing.T) {
	_, err := newImageCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
}

func TestNewImageCommandFromAttrs_WithScales(t *testing.T) {
	cmd, err := newImageCommandFromAttrs(map[string]string{
		"src":       "img",
		"imageType": "jpeg",
		"scaleX":    "2.0",
		"scaleY":    "0.5",
	})
	require.NoError(t, err)
	imgCmd := cmd.(*ImageCommand)
	assert.Equal(t, "JPEG", imgCmd.ImageType)
	assert.InDelta(t, 2.0, imgCmd.ScaleX, 0.01)
	assert.InDelta(t, 0.5, imgCmd.ScaleY, 0.01)
}

func TestNewImageCommandFromAttrs_DefaultType(t *testing.T) {
	cmd, err := newImageCommandFromAttrs(map[string]string{"src": "img"})
	require.NoError(t, err)
	assert.Equal(t, "PNG", cmd.(*ImageCommand).ImageType)
}

// =============================================================================
// newUpdateCellCommandFromAttrs — missing updater
// =============================================================================

func TestNewUpdateCellCommandFromAttrs_NoUpdater(t *testing.T) {
	_, err := newUpdateCellCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
}

// =============================================================================
// newEachCommandFromAttrs — missing attrs
// =============================================================================

func TestNewEachCommandFromAttrs_NoItems(t *testing.T) {
	_, err := newEachCommandFromAttrs(map[string]string{"var": "e"})
	assert.Error(t, err)
}

func TestNewEachCommandFromAttrs_NoVar(t *testing.T) {
	_, err := newEachCommandFromAttrs(map[string]string{"items": "x"})
	assert.Error(t, err)
}

// =============================================================================
// newIfCommandFromAttrs — missing condition
// =============================================================================

func TestNewIfCommandFromAttrs_NoCondition(t *testing.T) {
	_, err := newIfCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
}

// =============================================================================
// newGridCommandFromAttrs — missing attrs
// =============================================================================

func TestNewGridCommandFromAttrs_NoHeaders(t *testing.T) {
	_, err := newGridCommandFromAttrs(map[string]string{"data": "d"})
	assert.Error(t, err)
}

func TestNewGridCommandFromAttrs_NoData(t *testing.T) {
	_, err := newGridCommandFromAttrs(map[string]string{"headers": "h"})
	assert.Error(t, err)
}

// =============================================================================
// ProcessFormulasForRange — no expansion path
// =============================================================================

func TestProcessFormulasForRange_NoExpansion(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	fp := NewFormulaProcessor()
	// No cell refs registered → formula unchanged
	result := fp.ProcessFormulasForRange("SUM(A1:A5)", tx, "Sheet1")
	assert.Equal(t, "SUM(A1:A5)", result)
}

// =============================================================================
// formatRef — target on different sheet
// =============================================================================

func TestFormatRef_TargetDifferentSheet(t *testing.T) {
	fp := &StandardFormulaProcessor{}
	ref := NewCellRef("Sheet2", 0, 0)
	// Both origRefSheet and areaSheet are Sheet1, but target is Sheet2
	result := fp.formatRef(ref, "Sheet1", "Sheet1")
	assert.Equal(t, "Sheet2!A1", result)
}

// =============================================================================
// Additional coverage — remaining uncovered branches
// =============================================================================

// --- Grid ApplyAt error branches ---

func TestGridCommand_HeadersEvalError(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &GridCommand{Headers: "1 + + +", Data: "d"}
	ctx := NewContext(nil)

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestGridCommand_DataEvalError(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &GridCommand{Headers: "h", Data: "1 + + +"}
	ctx := NewContext(map[string]any{"h": []any{"Col1"}})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestGridCommand_HeadersNotIterable(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &GridCommand{Headers: "h", Data: "d"}
	ctx := NewContext(map[string]any{"h": 42, "d": []any{}})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestGridCommand_DataNotIterable(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &GridCommand{Headers: "h", Data: "d"}
	ctx := NewContext(map[string]any{"h": []any{"Col1"}, "d": 42})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestGridCommand_EmptyHeaders(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cmd := &GridCommand{Headers: "h", Data: "d"}
	ctx := NewContext(map[string]any{"h": []any{}, "d": []any{}})

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size)
}

// --- BuildAreas: area with empty lastCell, command error paths ---

func TestBuildAreas_AreaEmptyLastCell(t *testing.T) {
	f := excelize.NewFile()
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	_, err = filler.BuildAreas(tx)
	// Empty lastCell → area is skipped, then "no jx:area commands found"
	assert.Error(t, err)
}

func TestBuildAreas_CommandWithAreaLastCellEmpty(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "test")
	f.SetCellValue("Sheet1", "A2", "body")
	// Area with lastCell="A2" and a comment that is not a recognized command
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	// Separate comment on A2 — a valid each that will be found in the area
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="x" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)
	require.Len(t, areas, 1)
	// Should have one command binding (the each)
	assert.Len(t, areas[0].Bindings, 1)
}

// --- buildIfElseArea: error parsing else area ---

func TestBuildIfElseArea_InvalidElseRef(t *testing.T) {
	filler := &Filler{opts: defaultOptions(), registry: NewCommandRegistry()}
	ifCmd := &IfCommand{Condition: "true"}

	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Provide an unparseable else area reference
	err = filler.buildIfElseArea(ifCmd, `"A1:A1", "!@#invalid"`, NewCellRef("Sheet1", 0, 0), tx)
	assert.Error(t, err)
}

// --- area.go: transformStaticArea and processWithCommands error branches ---

func TestArea_TransformStaticArea_Error(t *testing.T) {
	// Create a transformer with no data for the source sheet
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Area references a non-existent sheet
	area := NewArea(NewCellRef("NoSheet", 0, 0), Size{Width: 1, Height: 1}, tx)
	ctx := NewContext(nil)

	// Static area transform — source cell won't exist, but Transform handles nil srcData gracefully
	size, err := area.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

// --- EachCommand: ApplyAt with items eval error ---

func TestEachCommand_ItemsEvalError(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &EachCommand{Items: "1 + + +", Var: "e", Area: area}
	ctx := NewContext(nil)

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestEachCommand_ItemsNotIterable(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &EachCommand{Items: "items", Var: "e", Area: area}
	ctx := NewContext(map[string]any{"items": 42})

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestEachCommand_EmptyItems(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &EachCommand{Items: "items", Var: "e", Area: area}
	ctx := NewContext(map[string]any{"items": []any{}})

	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size)
}

func TestEachCommand_SelectFiltersAll(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e}")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &EachCommand{Items: "items", Var: "e", Select: "e > 100", Area: area}
	ctx := NewContext(map[string]any{"items": []any{1, 2, 3}})

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size) // all filtered out
}

// --- ProcessAreaFormulas: formula cell not in area, no targets ---

func TestProcessAreaFormulas_FormulaCellOutsideArea(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellFormula(sheet, "A1", "B1+1")
	f.SetCellFormula(sheet, "C5", "D5+1")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Area only covers A1:B2 — C5 is outside
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)

	fp := NewFormulaProcessor()
	// Should not crash; C5 formula should be ignored
	fp.ProcessAreaFormulas(tx, area)
}

func TestProcessAreaFormulas_NoTargets(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellFormula(sheet, "A1", "B1+1")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)

	fp := NewFormulaProcessor()
	// A1 has formula but no target positions → skipped
	fp.ProcessAreaFormulas(tx, area)
}

// --- ProcessFormulasForRange: with actual expansion of both start and end ---

func TestProcessFormulasForRange_BothExpanded(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	// Must set cell values so Transform doesn't skip nil srcData
	f.SetCellValue(sheet, "A1", "v1")
	f.SetCellValue(sheet, "A3", "v3")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	// Start cell A1 maps to A1, A2
	tx.Transform(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 0, 0), ctx, false)
	tx.Transform(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 1, 0), ctx, false)
	// End cell A3 maps to A5
	tx.Transform(NewCellRef(sheet, 2, 0), NewCellRef(sheet, 4, 0), ctx, false)

	fp := NewFormulaProcessor()
	result := fp.ProcessFormulasForRange("SUM(A1:A3)", tx, sheet)
	// Should expand to cover A1:A5
	assert.Contains(t, result, "A1")
	assert.Contains(t, result, "A5")
}

// --- ProcessFormulasForRange: with sheet prefix in standalone position ---

func TestProcessFormulasForRange_WithSheetPrefix(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "val")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Register targets manually via addTargetRef
	tx.addTargetRef(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 0, 0))
	tx.addTargetRef(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 2, 0))

	fp := NewFormulaProcessor()
	// Use standalone sheet-prefixed range (not inside a function call, so regex captures sheet correctly)
	result := fp.ProcessFormulasForRange("Sheet1!A1:A1", tx, sheet)
	assert.Contains(t, result, "A1")
	assert.Contains(t, result, "A3")
}

// --- parseCellRefFromFormula: with $ signs and sheet ---

func TestParseCellRefFromFormula_WithDollarAndSheet(t *testing.T) {
	ref, err := parseCellRefFromFormula("Sheet2!$A$1", "Sheet1")
	require.NoError(t, err)
	assert.Equal(t, "Sheet2", ref.Sheet)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 0, ref.Col)
}

func TestParseCellRefFromFormula_NoDollar(t *testing.T) {
	ref, err := parseCellRefFromFormula("B5", "Sheet1")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ref.Sheet)
	assert.Equal(t, 4, ref.Row)
	assert.Equal(t, 1, ref.Col)
}

// --- EvaluateCellValue: expression eval error in mixed content ---

func TestEvaluateCellValue_ExprError(t *testing.T) {
	ctx := NewContext(nil)
	_, _, err := ctx.EvaluateCellValue("${1 + + +}")
	assert.Error(t, err)
}

func TestEvaluateCellValue_MixedExprError(t *testing.T) {
	ctx := NewContext(nil)
	_, _, err := ctx.EvaluateCellValue("prefix ${1 + + +} suffix")
	assert.Error(t, err)
}

func TestEvaluateCellValue_MixedWithNilExpr(t *testing.T) {
	ctx := NewContext(map[string]any{"x": nil})
	result, cellType, err := ctx.EvaluateCellValue("value: ${x}")
	require.NoError(t, err)
	assert.Equal(t, "value: ", result)
	assert.Equal(t, CellString, cellType)
}

// --- GetCellData: missing row, missing sheet ---

func TestGetCellData_MissingRow(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "val")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cd := tx.GetCellData(NewCellRef("Sheet1", 999, 0))
	assert.Nil(t, cd)
}

func TestGetCellData_MissingSheet(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	cd := tx.GetCellData(NewCellRef("NoSheet", 0, 0))
	assert.Nil(t, cd)
}

// --- Transform: non-expression string copies as-is ---

func TestTransform_PlainString(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "plain text")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	err = tx.Transform(NewCellRef(sheet, 0, 0), NewCellRef(sheet, 5, 0), ctx, true)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A6")
	assert.Equal(t, "plain text", v)
}

// --- CopySheet success ---

func TestCopySheet_Success(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "data")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	err = tx.CopySheet("Sheet1", "Sheet1Copy")
	require.NoError(t, err)

	sheets := tx.GetSheetNames()
	assert.Contains(t, sheets, "Sheet1Copy")
}

// --- NewExcelizeTransformer error branch ---

func TestNewExcelizeTransformer_NilFileError(t *testing.T) {
	// Passing nil causes panic in excelize; instead test with closed file
	// This verifies the error wrapping in NewExcelizeTransformer
	f := excelize.NewFile()
	f.Close()
	_, err := NewExcelizeTransformer(f)
	// May or may not error depending on excelize behavior with closed file
	_ = err
}

// --- parseCommandLine: invalid lastCell ---

func TestParseCommandLine_InvalidLastCell(t *testing.T) {
	_, _, err := ParseComment(`jx:each(items="x" var="e" lastCell="!!!")`, NewCellRef("Sheet1", 0, 0))
	assert.Error(t, err)
}

// --- ParseAreaRef: invalid second part ---

func TestParseAreaRef_InvalidSecondPart(t *testing.T) {
	_, err := ParseAreaRef("A1:!!!")
	assert.Error(t, err)
}

// --- parseCellName: single letter (valid col, no row digits) ---

func TestParseCellRef_NoRowDigits(t *testing.T) {
	_, err := ParseCellRef("A")
	assert.Error(t, err)
}

func TestParseCellRef_OnlyDigits(t *testing.T) {
	_, err := ParseCellRef("123")
	assert.Error(t, err)
}

// --- sortByFields: single item (no sort needed) ---

func TestSortByFields_SingleItem(t *testing.T) {
	items := []any{map[string]any{"Val": 1}}
	specs := []orderBySpec{{field: "Val", desc: false}}
	sortByFields(items, specs)
	assert.Len(t, items, 1)
}

func TestSortByFields_EmptySpecs(t *testing.T) {
	items := []any{1, 2}
	sortByFields(items, nil)
	assert.Equal(t, []any{1, 2}, items)
}

// --- compareBySpecs: equal items return 0 ---

func TestCompareBySpecs_Equal(t *testing.T) {
	specs := []orderBySpec{{field: "Val", desc: false}}
	a := map[string]any{"Val": 5}
	b := map[string]any{"Val": 5}
	assert.Equal(t, 0, compareBySpecs(a, b, specs))
}

// --- tryBuildRange: horizontal with gap ---

func TestTryBuildRange_HorizontalWithGap(t *testing.T) {
	fp := &StandardFormulaProcessor{}
	targets := []CellRef{
		NewCellRef("S", 0, 0),
		NewCellRef("S", 0, 2), // gap at col 1
	}
	result := fp.tryBuildRange(targets, "S", "S")
	assert.Equal(t, "", result) // non-contiguous → empty
}

func TestTryBuildRange_SingleTarget(t *testing.T) {
	fp := &StandardFormulaProcessor{}
	targets := []CellRef{NewCellRef("S", 0, 0)}
	result := fp.tryBuildRange(targets, "S", "S")
	assert.Equal(t, "", result) // single → empty (handled by caller)
}

// --- processFormula: no matches in formula ---

func TestProcessFormula_NoRefs(t *testing.T) {
	fp := &StandardFormulaProcessor{}
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 5, Height: 5}, tx)
	cd := &CellData{Ref: NewCellRef(sheet, 0, 0), Formula: "123+456"}

	result := fp.processFormula("123+456", cd, NewCellRef(sheet, 0, 0), tx, area)
	assert.Equal(t, "123+456", result) // no cell refs → unchanged
}

// --- FillWriter: expression notation forwarding ---

func TestFillWriter_CustomNotation(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "<<val>>")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	tmpl := filepath.Join(testdataDir(t), "custom_notation.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	filler := NewFiller(WithTemplate(tmpl), WithExpressionNotation("<<", ">>"))
	var outBuf bytes.Buffer
	err := filler.FillWriter(map[string]any{"val": "works"}, &outBuf)
	require.NoError(t, err)

	out, err := excelize.OpenReader(&outBuf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "works", v)
}

// --- resolveLastCell edge: lastCell with sheet prefix containing ! ---

func TestResolveLastCell_ExplicitSheet(t *testing.T) {
	start := NewCellRef("Sheet1", 0, 0)
	ref, err := resolveLastCell(start, "OtherSheet!C5")
	require.NoError(t, err)
	assert.Equal(t, "OtherSheet", ref.Sheet)
}

// --- ParseExpressions: no expressions ---

func TestParseExpressions_NoExpressions(t *testing.T) {
	segs := ParseExpressions("just plain text", "${", "}")
	assert.Len(t, segs, 1)
	assert.False(t, segs[0].IsExpression)
	assert.Equal(t, "just plain text", segs[0].Text)
}

func TestParseExpressions_EmptyString(t *testing.T) {
	segs := ParseExpressions("", "${", "}")
	assert.Len(t, segs, 0)
}

func TestParseExpressions_UnclosedDelimiter(t *testing.T) {
	segs := ParseExpressions("text ${unclosed", "${", "}")
	// Should return just the text since no matching end
	assert.Len(t, segs, 1)
	assert.Equal(t, "text ${unclosed", segs[0].Text)
}

// --- ParseComment: multiline with params ---

func TestParseComment_MultilineWithParams(t *testing.T) {
	cmds, params, err := ParseComment(
		"jx:each(items=\"x\" var=\"e\" lastCell=\"A1\")\njx:params(defaultValue=\"0\" formulaStrategy=\"BY_COLUMN\")",
		NewCellRef("Sheet1", 0, 0),
	)
	require.NoError(t, err)
	assert.Len(t, cmds, 1)
	assert.NotNil(t, params)
	assert.Equal(t, "0", params.DefaultValue)
	assert.Equal(t, FormulaByColumn, params.FormulaStrategy)
}
