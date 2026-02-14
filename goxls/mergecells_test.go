package goxls

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestMergeCellsCommand_Basic(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Merged Header")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	cmd := &MergeCellsCommand{Cols: "3", Rows: "1"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	merges, _ := out.GetMergeCells(sheet)
	assert.True(t, len(merges) > 0, "should have merged cells")
}

func TestMergeCellsCommand_Dynamic(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"numCols": 4, "numRows": 2})
	cmd := &MergeCellsCommand{Cols: "numCols", Rows: "numRows"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 4, Height: 2}, size)
}

func TestMergeCellsCommand_MinThreshold(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	// Cols=2 but minCols=3 → skip merge
	cmd := &MergeCellsCommand{Cols: "2", Rows: "1", MinCols: "3"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size)

	// Should NOT have merged
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	merges, _ := out.GetMergeCells(sheet)
	assert.Len(t, merges, 0)
}

func TestMergeCellsCommand_SingleCell(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	cmd := &MergeCellsCommand{Cols: "1", Rows: "1"}
	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size) // no actual merge
}

func TestNewMergeCellsCommandFromAttrs(t *testing.T) {
	cmd, err := newMergeCellsCommandFromAttrs(map[string]string{
		"cols": "5", "rows": "3", "minCols": "2",
	})
	require.NoError(t, err)
	mc := cmd.(*MergeCellsCommand)
	assert.Equal(t, "5", mc.Cols)
	assert.Equal(t, "3", mc.Rows)
	assert.Equal(t, "2", mc.MinCols)
}
