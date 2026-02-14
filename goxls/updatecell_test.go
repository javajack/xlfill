package goxls

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// testUpdater implements CellDataUpdater for testing.
type testUpdater struct {
	called  int
	lastRef CellRef
}

func (u *testUpdater) UpdateCellData(cd *CellData, targetCell CellRef, ctx *Context) {
	u.called++
	u.lastRef = targetCell
	cd.Value = "UPDATED"
}

// formulaUpdater sets a formula on the cell.
type formulaUpdater struct{}

func (u *formulaUpdater) UpdateCellData(cd *CellData, targetCell CellRef, ctx *Context) {
	cd.Formula = "SUM(A1:A10)"
	cd.Value = nil
}

func TestUpdateCellCommand_BasicUpdate(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Original")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	updater := &testUpdater{}
	ctx := NewContext(map[string]any{"myUpdater": updater})

	cmd := &UpdateCellCommand{Updater: "myUpdater"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
	assert.Equal(t, 1, updater.called)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "UPDATED", v)
}

func TestUpdateCellCommand_FormulaUpdate(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"myUpdater": &formulaUpdater{}})
	cmd := &UpdateCellCommand{Updater: "myUpdater"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "A1")
	assert.Equal(t, "SUM(A1:A10)", formula)
}

func TestUpdateCellCommand_NilUpdater(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	cmd := &UpdateCellCommand{Updater: "missing"}
	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "not found")
}

func TestUpdateCellCommand_WrongType(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"myUpdater": "not-an-updater"})
	cmd := &UpdateCellCommand{Updater: "myUpdater"}
	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "CellDataUpdater")
}

func TestNewUpdateCellCommandFromAttrs(t *testing.T) {
	cmd, err := newUpdateCellCommandFromAttrs(map[string]string{"updater": "myUp"})
	require.NoError(t, err)
	assert.Equal(t, "updateCell", cmd.Name())
	uc := cmd.(*UpdateCellCommand)
	assert.Equal(t, "myUp", uc.Updater)
}

func TestNewUpdateCellCommandFromAttrs_MissingUpdater(t *testing.T) {
	_, err := newUpdateCellCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
}
