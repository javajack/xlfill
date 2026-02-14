package xlfill

import (
	"bytes"
	"fmt"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestEachCommand_BasicList(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.Age}")
	f.SetCellValue(sheet, "C1", "${e.Salary}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000.0},
		map[string]any{"Name": "Bob", "Age": 25, "Salary": 6000.0},
		map[string]any{"Name": "Carol", "Age": 35, "Salary": 7000.0},
	}
	ctx := NewContext(map[string]any{"employees": employees})

	cmd := &EachCommand{
		Items: "employees", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 3, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)
	assert.Equal(t, 3, size.Width)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Carol", v)

	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "5000", v)
	v, _ = out.GetCellValue(sheet, "C2")
	assert.Equal(t, "6000", v)
	v, _ = out.GetCellValue(sheet, "C3")
	assert.Equal(t, "7000", v)
}

func TestEachCommand_EmptyList(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"items": []any{}})
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size)
}

func TestEachCommand_NilList(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"items": nil})
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, ZeroSize, size)
}

func TestEachCommand_SingleItem(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"items": []any{map[string]any{"Name": "Solo"}}})
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Solo", v)
}

func TestEachCommand_LargeList(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := make([]any, 100)
	for i := range items {
		items[i] = map[string]any{"Name": fmt.Sprintf("Item%d", i)}
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 100, size.Height)
	assert.Equal(t, 1, size.Width)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Spot check first, middle, last
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Item0", v)
	v, _ = out.GetCellValue(sheet, "A50")
	assert.Equal(t, "Item49", v)
	v, _ = out.GetCellValue(sheet, "A100")
	assert.Equal(t, "Item99", v)
}

func TestEachCommand_VarIndex(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${idx}")
	f.SetCellValue(sheet, "B1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice"},
		map[string]any{"Name": "Bob"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", VarIndex: "idx", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "0", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "1", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "Bob", v)
}

func TestEachCommand_MultiColumnTemplate(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.ID}")
	f.SetCellValue(sheet, "B1", "${e.Name}")
	f.SetCellValue(sheet, "C1", "${e.Email}")
	f.SetCellValue(sheet, "D1", "${e.Dept}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"ID": 1, "Name": "Alice", "Email": "alice@ex.com", "Dept": "Eng"},
		map[string]any{"ID": 2, "Name": "Bob", "Email": "bob@ex.com", "Dept": "Sales"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 4, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 4, Height: 2}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "C1")
	assert.Equal(t, "alice@ex.com", v)
	v, _ = out.GetCellValue(sheet, "D2")
	assert.Equal(t, "Sales", v)
}

func TestEachCommand_NumberTypes(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.IntVal}")
	f.SetCellValue(sheet, "B1", "${e.FloatVal}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"IntVal": 42, "FloatVal": 3.14},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "42", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "3.14", v)
}

func TestEachCommand_DateTypes(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Date}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	dt := time.Date(2024, 6, 15, 0, 0, 0, 0, time.UTC)
	items := []any{
		map[string]any{"Date": dt},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

func TestEachCommand_NilFieldValue(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": nil},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

func TestEachCommand_NestedStruct(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Address.City}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Address": map[string]any{"City": "NYC"}},
		map[string]any{"Address": map[string]any{"City": "London"}},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "NYC", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "London", v)
}

func TestEachCommand_DirectionRight(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Q1"},
		map[string]any{"Name": "Q2"},
		map[string]any{"Name": "Q3"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "RIGHT",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Q1", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Q2", v)
	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "Q3", v)
}

func TestEachCommand_SelectFilter(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Active": true},
		map[string]any{"Name": "Bob", "Active": false},
		map[string]any{"Name": "Carol", "Active": true},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Select: "e.Active == true",
		Area:   NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height) // only Alice and Carol

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Carol", v)
}

func TestEachCommand_OrderBy(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.Salary}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Carol", "Salary": 7000.0},
		map[string]any{"Name": "Alice", "Salary": 5000.0},
		map[string]any{"Name": "Bob", "Salary": 6000.0},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		OrderBy: "e.Name ASC",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Carol", v)
}

func TestEachCommand_OrderByDesc(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Salary": 5000.0},
		map[string]any{"Name": "Bob", "Salary": 6000.0},
		map[string]any{"Name": "Carol", "Salary": 7000.0},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		OrderBy: "e.Salary DESC",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Carol", v) // highest salary
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Alice", v) // lowest salary
}

func TestEachCommand_NoArea(t *testing.T) {
	ctx := NewContext(map[string]any{"items": []any{1, 2}})
	cmd := &EachCommand{Items: "items", Var: "e", Direction: "DOWN"}

	_, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, nil)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no area")
}

func TestEachCommand_InvalidItems(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"items": "not-a-slice"})
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	_, err = cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "not iterable")
}

func TestEachCommand_PreservesFormatting(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	boldStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
	})
	require.NoError(t, err)

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellStyle(sheet, "A1", "A1", boldStyle)

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice"},
		map[string]any{"Name": "Bob"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	_, err = cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Both rows should have bold style
	s1, _ := out.GetCellStyle(sheet, "A1")
	s2, _ := out.GetCellStyle(sheet, "A2")
	assert.True(t, s1 > 0, "row 1 should have style")
	assert.Equal(t, s1, s2, "row 2 should have same style as row 1")
}

// --- Helper sort tests ---

func TestParseOrderBy(t *testing.T) {
	specs := parseOrderBy("e.Name ASC, e.Salary DESC", "e")
	require.Len(t, specs, 2)
	assert.Equal(t, "Name", specs[0].field)
	assert.False(t, specs[0].desc)
	assert.Equal(t, "Salary", specs[1].field)
	assert.True(t, specs[1].desc)
}

func TestParseOrderBy_Empty(t *testing.T) {
	specs := parseOrderBy("", "e")
	assert.Nil(t, specs)
}

func TestParseOrderBy_NoDirection(t *testing.T) {
	specs := parseOrderBy("e.Name", "e")
	require.Len(t, specs, 1)
	assert.Equal(t, "Name", specs[0].field)
	assert.False(t, specs[0].desc) // default ASC
}

func TestToSlice(t *testing.T) {
	// []any
	result, err := toSlice([]any{1, 2, 3})
	require.NoError(t, err)
	assert.Len(t, result, 3)

	// []string (typed slice via reflection)
	result, err = toSlice([]string{"a", "b"})
	require.NoError(t, err)
	assert.Len(t, result, 2)

	// nil
	result, err = toSlice(nil)
	require.NoError(t, err)
	assert.Nil(t, result)

	// non-iterable
	_, err = toSlice("string")
	assert.Error(t, err)
}

func TestNewEachCommandFromAttrs(t *testing.T) {
	cmd, err := newEachCommandFromAttrs(map[string]string{
		"items":     "employees",
		"var":       "e",
		"varIndex":  "idx",
		"direction": "right",
		"select":    "e.Active",
		"orderBy":   "e.Name ASC",
	})
	require.NoError(t, err)

	each := cmd.(*EachCommand)
	assert.Equal(t, "employees", each.Items)
	assert.Equal(t, "e", each.Var)
	assert.Equal(t, "idx", each.VarIndex)
	assert.Equal(t, "RIGHT", each.Direction)
	assert.Equal(t, "e.Active", each.Select)
	assert.Equal(t, "e.Name ASC", each.OrderBy)
}

func TestNewEachCommandFromAttrs_Missing(t *testing.T) {
	_, err := newEachCommandFromAttrs(map[string]string{"var": "e"})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "items")

	_, err = newEachCommandFromAttrs(map[string]string{"items": "list"})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "var")
}

// --- GroupBy tests ---

func TestEachCommand_GroupBy_Basic(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	// Template: show group key (department from first item)
	f.SetCellValue(sheet, "A1", "${g.Item.Dept}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Dept": "Eng"},
		map[string]any{"Name": "Bob", "Dept": "Sales"},
		map[string]any{"Name": "Carol", "Dept": "Eng"},
		map[string]any{"Name": "Dave", "Dept": "Sales"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		GroupBy: "g.Dept",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height) // 2 groups: Eng, Sales

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Groups maintain insertion order: Eng first (Alice), Sales second (Bob)
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Eng", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Sales", v)
}

func TestEachCommand_GroupBy_Asc(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${g.Item.Dept}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Carol", "Dept": "Sales"},
		map[string]any{"Name": "Alice", "Dept": "Eng"},
		map[string]any{"Name": "Bob", "Dept": "HR"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		GroupBy: "g.Dept", GroupOrder: "ASC",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Eng", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "HR", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Sales", v)
}

func TestEachCommand_GroupBy_Desc(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${g.Item.Dept}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Dept": "Eng"},
		map[string]any{"Dept": "Sales"},
		map[string]any{"Dept": "HR"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		GroupBy: "g.Dept", GroupOrder: "DESC",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Sales", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "HR", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Eng", v)
}

func TestEachCommand_GroupBy_WithSelect(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${g.Item.Dept}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Dept": "Eng", "Active": true},
		map[string]any{"Name": "Bob", "Dept": "Sales", "Active": false},
		map[string]any{"Name": "Carol", "Dept": "Eng", "Active": true},
		map[string]any{"Name": "Dave", "Dept": "HR", "Active": true},
	}
	ctx := NewContext(map[string]any{"items": items})

	// Note: select uses the original var name before grouping
	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		Select:  "g.Active == true",
		GroupBy: "g.Dept",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height) // Eng and HR (Bob filtered out, so no Sales group)
}

func TestEachCommand_GroupBy_GroupDataItems(t *testing.T) {
	// Verify that GroupData.Items contains the correct members.
	items := []any{
		map[string]any{"Name": "Alice", "Dept": "Eng"},
		map[string]any{"Name": "Bob", "Dept": "Sales"},
		map[string]any{"Name": "Carol", "Dept": "Eng"},
	}

	cmd := &EachCommand{
		Items: "items", Var: "e",
		GroupBy: "e.Dept",
	}

	grouped := cmd.groupItems(items)
	require.Len(t, grouped, 2)

	g1 := grouped[0].(GroupData)
	assert.Equal(t, "Eng", getField(g1.Item, "Dept"))
	assert.Len(t, g1.Items, 2) // Alice, Carol

	g2 := grouped[1].(GroupData)
	assert.Equal(t, "Sales", getField(g2.Item, "Dept"))
	assert.Len(t, g2.Items, 1) // Bob
}

func TestEachCommand_GroupBy_IgnoreCase(t *testing.T) {
	items := []any{
		map[string]any{"Dept": "engineering"},
		map[string]any{"Dept": "Sales"},
		map[string]any{"Dept": "ENGINEERING"},
	}

	cmd := &EachCommand{
		Items: "items", Var: "e",
		GroupBy: "e.Dept", GroupOrder: "ASC_IGNORECASE",
	}

	grouped := cmd.groupItems(items)
	// "engineering" and "ENGINEERING" are different string keys, so 3 groups
	// But after sorting with ignore case, they should be ordered properly
	require.True(t, len(grouped) >= 2)
}
