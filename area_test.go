package xlfill

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// createAreaTestTemplate creates a simple template for area processing tests.
// Layout (Sheet1):
//
//	A1: "Name"           B1: "Salary"
//	A2: "${e.Name}"      B2: "${e.Salary}"
func createAreaTestTemplate(t *testing.T) *ExcelizeTransformer {
	t.Helper()
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Salary")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Salary}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	return tx
}

func TestArea_ApplyAt_StaticCells(t *testing.T) {
	// Area with no commands — just transforms all cells as-is.
	tx := createAreaTestTemplate(t)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"e": map[string]any{"Name": "Alice", "Salary": 5000.0},
	})

	area := NewArea(
		NewCellRef("Sheet1", 0, 0), // A1
		Size{Width: 2, Height: 2},
		tx,
	)

	size, err := area.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 2}, size)

	// Verify the expression cell was evaluated
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))

	f, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v)
	v, _ = f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "5000", v)

	// Static header preserved
	v, _ = f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "Name", v)
}

func TestArea_ApplyAt_NilTransformer(t *testing.T) {
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 1, Height: 1}, nil)
	ctx := NewContext(nil)
	_, err := area.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no transformer")
}

func TestArea_ApplyAt_SingleCommand(t *testing.T) {
	// Area with a jx:each that repeats 3 rows.
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Header row (static)
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Salary")

	// Data row (template)
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Salary}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Alice", "Salary": 5000.0},
		map[string]any{"Name": "Bob", "Salary": 6000.0},
		map[string]any{"Name": "Carol", "Salary": 7000.0},
	}

	ctx := NewContext(map[string]any{"employees": employees})

	// Build area: A1:B2 (2 wide, 2 high)
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)

	// EachCommand on row 2 (A2:B2), 1 row high
	eachCmd := &EachCommand{
		Items: "employees",
		Var:   "e",
		Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1}, tx),
	}

	area.AddCommand(eachCmd, NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1})

	size, err := area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// 1 header row + 3 data rows = 4 total height
	assert.Equal(t, 4, size.Height)
	assert.Equal(t, 2, size.Width)

	// Verify output
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))

	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Name", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Carol", v)

	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "5000", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "6000", v)
	v, _ = out.GetCellValue(sheet, "B4")
	assert.Equal(t, "7000", v)
}

func TestArea_ApplyAt_MultipleCommands(t *testing.T) {
	// Two sequential each commands in one area.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Employees:")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "A3", "Departments:")
	f.SetCellValue(sheet, "A4", "${d.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"employees":   []any{map[string]any{"Name": "Alice"}, map[string]any{"Name": "Bob"}},
		"departments": []any{map[string]any{"Name": "Engineering"}, map[string]any{"Name": "Sales"}},
	})

	// Area A1:A4 (1 wide, 4 high)
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 4}, tx)

	// First each at A2 (row 1)
	each1 := &EachCommand{
		Items: "employees", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx),
	}
	area.AddCommand(each1, NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1})

	// Second each at A4 (row 3)
	each2 := &EachCommand{
		Items: "departments", Var: "d", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 3, 0), Size{Width: 1, Height: 1}, tx),
	}
	area.AddCommand(each2, NewCellRef(sheet, 3, 0), Size{Width: 1, Height: 1})

	size, err := area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// Row layout: "Employees:" (1) + 2 employees + "Departments:" (1) + 2 departments = 6
	assert.Equal(t, 6, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))

	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Employees:", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Departments:", v)
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "Engineering", v)
	v, _ = out.GetCellValue(sheet, "A6")
	assert.Equal(t, "Sales", v)
}

func TestArea_ApplyAt_CommandContraction(t *testing.T) {
	// Empty list → zero size from command.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "A3", "Footer")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"employees": []any{}})

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 3}, tx)

	each := &EachCommand{
		Items: "employees", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx),
	}
	area.AddCommand(each, NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1})

	size, err := area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// Header (1) + 0 from empty each + Footer (1) = 2
	assert.Equal(t, 2, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Footer", v)
}

func TestArea_ApplyAt_CommandExpansion(t *testing.T) {
	// 5-item list pushes footer down.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "A3", "Footer")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := make([]any, 5)
	for i := range items {
		items[i] = map[string]any{"Name": fmt.Sprintf("Item%d", i)}
	}
	ctx := NewContext(map[string]any{"items": items})

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 3}, tx)
	each := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx),
	}
	area.AddCommand(each, NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1})

	size, err := area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// Header(1) + 5 items + Footer(1) = 7
	assert.Equal(t, 7, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
	v, _ = out.GetCellValue(sheet, "A7")
	assert.Equal(t, "Footer", v)

	for i := 0; i < 5; i++ {
		v, _ := out.GetCellValue(sheet, fmt.Sprintf("A%d", i+2))
		assert.Equal(t, fmt.Sprintf("Item%d", i), v)
	}
}

func TestArea_ClearCells(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Hello")
	f.SetCellValue(sheet, "B1", "World")
	f.SetCellValue(sheet, "A2", "Foo")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)
	area.ClearCells()

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "", v)
}

func TestArea_ApplyAt_NestedCommands(t *testing.T) {
	// Each with If inside — conditional rendering per row.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "HIGH")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Alice", "Salary": 8000.0},
		map[string]any{"Name": "Bob", "Salary": 3000.0},
		map[string]any{"Name": "Carol", "Salary": 9000.0},
	}
	ctx := NewContext(map[string]any{"employees": employees})

	// Inner if area (just B1 cell)
	ifArea := NewArea(NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1}, tx)
	ifCmd := &IfCommand{
		Condition: "e.Salary > 5000",
		IfArea:    ifArea,
	}

	// Each command's area is A1:B1 with an if command on B1
	eachArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	eachArea.AddCommand(ifCmd, NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1})

	eachCmd := &EachCommand{
		Items: "employees", Var: "e", Direction: "DOWN",
		Area: eachArea,
	}

	// Root area is just the each
	rootArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	rootArea.AddCommand(eachCmd, NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1})

	size, err := rootArea.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Alice (high salary) — should have "HIGH" in B
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "HIGH", v)

	// Bob (low salary) — B2 should be empty (no else area)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)

	// Carol (high salary)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Carol", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "HIGH", v)
}

func TestArea_ApplyAt_TargetOffset(t *testing.T) {
	// Apply area at a different target position than source.
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Hello")
	f.SetCellValue(sheet, "B1", "World")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)

	// Apply at row 5 (0-based=4), col C (0-based=2)
	size, err := area.ApplyAt(NewCellRef(sheet, 4, 2), ctx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "C5")
	assert.Equal(t, "Hello", v)
	v, _ = out.GetCellValue(sheet, "D5")
	assert.Equal(t, "World", v)
}
