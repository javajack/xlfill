package xlfill

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestIfCommand_True(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Visible")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"show": true})

	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "show", IfArea: ifArea}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Visible", v)
}

func TestIfCommand_False(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "If Content")
	f.SetCellValue(sheet, "A2", "Else Content")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"show": false})

	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	elseArea := NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "show", IfArea: ifArea, ElseArea: elseArea}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Else area content written to target position (A1)
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Else Content", v)
}

func TestIfCommand_FalseNoElse(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Hidden")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"show": false})

	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "show", IfArea: ifArea}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size)
}

func TestIfCommand_ExpressionCondition(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Big")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"amount": 5000})

	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "amount > 1000", IfArea: ifArea}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

func TestIfCommand_InsideEach(t *testing.T) {
	// If inside each loop — different conditions per iteration.
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "Premium")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Alice", "Salary": 8000.0},
		map[string]any{"Name": "Bob", "Salary": 3000.0},
		map[string]any{"Name": "Carol", "Salary": 9000.0},
	}
	ctx := NewContext(map[string]any{"employees": employees})

	// If area (just B1)
	ifArea := NewArea(NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1}, tx)
	ifCmd := &IfCommand{Condition: "e.Salary > 5000", IfArea: ifArea}

	// Each area = A1:B1 with if on B1
	eachArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	eachArea.AddCommand(ifCmd, NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1})

	eachCmd := &EachCommand{
		Items: "employees", Var: "e", Direction: "DOWN",
		Area: eachArea,
	}

	size, err := eachCmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Premium", v) // Alice salary > 5000

	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)
	// B2 should not have "Premium" (Bob salary < 5000)

	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Carol", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "Premium", v) // Carol salary > 5000
}

func TestIfCommand_PreservesFormatting(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	boldStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
	})
	require.NoError(t, err)

	f.SetCellValue(sheet, "A1", "Styled")
	f.SetCellStyle(sheet, "A1", "A1", boldStyle)

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"show": true})
	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "show", IfArea: ifArea}

	// Apply to row 3
	size, err := cmd.ApplyAt(NewCellRef(sheet, 2, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	s, _ := out.GetCellStyle(sheet, "A3")
	assert.True(t, s > 0, "target cell should have style preserved")
}

func TestIfCommand_InvalidCondition(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	ifArea := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "???invalid!!!", IfArea: ifArea}

	_, err = cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	assert.Error(t, err)
}

func TestNewIfCommandFromAttrs(t *testing.T) {
	cmd, err := newIfCommandFromAttrs(map[string]string{"condition": "x > 5"})
	require.NoError(t, err)
	assert.Equal(t, "if", cmd.Name())

	ifCmd := cmd.(*IfCommand)
	assert.Equal(t, "x > 5", ifCmd.Condition)
}

func TestNewIfCommandFromAttrs_MissingCondition(t *testing.T) {
	_, err := newIfCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "condition")
}

func TestIfCommand_WithElseExpression(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}: Premium")
	f.SetCellValue(sheet, "A2", "${e.Name}: Standard")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"e":     map[string]any{"Name": "Alice", "VIP": false},
		"items": []any{map[string]any{"Name": "Alice", "VIP": false}},
	})

	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	elseArea := NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{Condition: "e.VIP == true", IfArea: ifArea, ElseArea: elseArea}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 4, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A5")
	assert.Equal(t, "Alice: Standard", v)
}
