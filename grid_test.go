package xlfill

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestGridCommand_BasicGrid(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"headers": []any{"Name", "Age", "City"},
		"data": []any{
			[]any{"Alice", 30, "NYC"},
			[]any{"Bob", 25, "London"},
		},
	})

	cmd := &GridCommand{Headers: "headers", Data: "data"}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 3}, size) // 1 header + 2 data rows

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Name", v)
	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "City", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "C3")
	assert.Equal(t, "London", v)
}

func TestGridCommand_NilHeaders(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"headers": nil, "data": []any{}})
	cmd := &GridCommand{Headers: "headers", Data: "data"}
	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size)
}

func TestGridCommand_NilData(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"headers": []any{"H1", "H2"},
		"data":    nil,
	})
	cmd := &GridCommand{Headers: "headers", Data: "data"}
	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size) // headers only
}

func TestGridCommand_ObjectDataWithProps(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"headers": []any{"Name", "Salary"},
		"data": []any{
			map[string]any{"Name": "Alice", "Salary": 5000, "Hidden": "x"},
			map[string]any{"Name": "Bob", "Salary": 6000, "Hidden": "y"},
		},
	})

	cmd := &GridCommand{Headers: "headers", Data: "data", Props: "Name, Salary"}
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
	assert.Equal(t, "6000", v)
}

func TestNewGridCommandFromAttrs(t *testing.T) {
	cmd, err := newGridCommandFromAttrs(map[string]string{
		"headers": "h", "data": "d", "props": "A,B",
	})
	require.NoError(t, err)
	g := cmd.(*GridCommand)
	assert.Equal(t, "h", g.Headers)
	assert.Equal(t, "d", g.Data)
	assert.Equal(t, "A,B", g.Props)
}

func TestNewGridCommandFromAttrs_MissingHeaders(t *testing.T) {
	_, err := newGridCommandFromAttrs(map[string]string{"data": "d"})
	assert.Error(t, err)
}

func TestNewGridCommandFromAttrs_MissingData(t *testing.T) {
	_, err := newGridCommandFromAttrs(map[string]string{"headers": "h"})
	assert.Error(t, err)
}
