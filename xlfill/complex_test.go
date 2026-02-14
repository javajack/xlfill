package xlfill

import (
	"bytes"
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// --- Phase 16.1: Nested Each ---

func TestNested_EachInsideEach(t *testing.T) {
	// Departments with employees — manually wired nested each.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${d.Name}")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Salary}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	departments := []any{
		map[string]any{
			"Name": "Engineering",
			"Employees": []any{
				map[string]any{"Name": "Alice", "Salary": 5000},
				map[string]any{"Name": "Bob", "Salary": 6000},
			},
		},
		map[string]any{
			"Name": "Sales",
			"Employees": []any{
				map[string]any{"Name": "Carol", "Salary": 4000},
			},
		},
	}
	ctx := NewContext(map[string]any{"departments": departments})

	// Inner each: iterates d.Employees on row 2 (A2:B2)
	innerEach := &EachCommand{
		Items: "d.Employees", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1}, tx),
	}

	// Outer each area: A1:B2 (2 rows: dept header + employee row)
	outerArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)
	outerArea.AddCommand(innerEach, NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1})

	outerEach := &EachCommand{
		Items: "departments", Var: "d", Direction: "DOWN",
		Area: outerArea,
	}

	size, err := outerEach.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)

	// Engineering: 1 header + 2 employees = 3 rows
	// Sales: 1 header + 1 employee = 2 rows
	// Total = 5
	assert.Equal(t, 5, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Engineering", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "5000", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Sales", v)
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "Carol", v)
}

// --- Phase 16.2: Each + If Combined ---

func TestCombined_EachWithIf(t *testing.T) {
	// Each with if — some rows have conditional content.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "VIP")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "VIP": true},
		map[string]any{"Name": "Bob", "VIP": false},
		map[string]any{"Name": "Carol", "VIP": true},
	}
	ctx := NewContext(map[string]any{"items": items})

	ifArea := NewArea(NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1}, tx)
	ifCmd := &IfCommand{Condition: "e.VIP == true", IfArea: ifArea}

	eachArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	eachArea.AddCommand(ifCmd, NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1})

	eachCmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
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

	v, _ := out.GetCellValue(sheet, "B1")
	assert.Equal(t, "VIP", v) // Alice is VIP
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "VIP", v) // Carol is VIP
}

func TestCombined_EachWithIfElse(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "Active")
	f.SetCellValue(sheet, "B2", "Inactive")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Active": true},
		map[string]any{"Name": "Bob", "Active": false},
	}
	ctx := NewContext(map[string]any{"items": items})

	ifArea := NewArea(NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1}, tx)
	elseArea := NewArea(NewCellRef(sheet, 1, 1), Size{Width: 1, Height: 1}, tx)
	ifCmd := &IfCommand{Condition: "e.Active == true", IfArea: ifArea, ElseArea: elseArea}

	eachArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	eachArea.AddCommand(ifCmd, NewCellRef(sheet, 0, 1), Size{Width: 1, Height: 1})

	eachCmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: eachArea,
	}

	size, err := eachCmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Active", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "Inactive", v)
}

// --- Phase 16.4: Edge Cases ---

func TestEdge_LargeDataSet(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "ID")
	f.SetCellValue(sheet, "B1", "Value")
	f.SetCellValue(sheet, "A2", "${e.ID}")
	f.SetCellValue(sheet, "B2", "${e.Value}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="items" var="e" lastCell="B2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "large_data.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	items := make([]any, 1000)
	for i := range items {
		items[i] = map[string]any{"ID": i + 1, "Value": fmt.Sprintf("Val%d", i)}
	}

	out, err := FillBytes(tmpl, map[string]any{"items": items})
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	// Spot check
	v, _ := outFile.GetCellValue(sheet, "A2")
	assert.Equal(t, "1", v)
	v, _ = outFile.GetCellValue(sheet, "A1001")
	assert.Equal(t, "1000", v)
}

func TestEdge_SpecialCharInData(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": `<script>alert("xss")</script>`},
		map[string]any{"Name": `Tom & Jerry "quotes" 'apostrophe'`},
		map[string]any{"Name": `Line1\nLine2`},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
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
	assert.Contains(t, v, "alert")
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Contains(t, v, "Tom & Jerry")
}

func TestEdge_UnicodeData(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "日本語テスト"},
		map[string]any{"Name": "中文测试"},
		map[string]any{"Name": "Ñoño España"},
		map[string]any{"Name": "Привет мир"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 4, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "日本語テスト", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Привет мир", v)
}

func TestEdge_LongString(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${val}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	longStr := strings.Repeat("A", 32767) // Excel max cell length
	ctx := NewContext(map[string]any{"val": longStr})

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	_, err = area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Len(t, v, 32767)
}

// --- Phase 17: Custom Commands & Extensibility ---

// customSumCommand is a test custom command that sums values.
type customSumCommand struct {
	Expr string
}

func (c *customSumCommand) Name() string { return "customSum" }
func (c *customSumCommand) Reset()       {}
func (c *customSumCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	val, err := ctx.Evaluate(c.Expr)
	if err != nil {
		return ZeroSize, err
	}
	transformer.SetCellValue(cellRef, val)
	return Size{Width: 1, Height: 1}, nil
}

func TestCustomCommand_Registration(t *testing.T) {
	reg := NewCommandRegistry()
	reg.Register("customSum", func(attrs map[string]string) (Command, error) {
		return &customSumCommand{Expr: attrs["expr"]}, nil
	})

	cmd, err := reg.Create("customSum", map[string]string{"expr": "1+2"})
	require.NoError(t, err)
	require.NotNil(t, cmd)
	assert.Equal(t, "customSum", cmd.Name())
}

func TestCustomCommand_UnknownIgnored(t *testing.T) {
	reg := NewCommandRegistry()
	cmd, err := reg.Create("unknownCmd", map[string]string{})
	require.NoError(t, err)
	assert.Nil(t, cmd)
}

func TestCustomCommand_WithFiller(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Result")
	f.SetCellValue(sheet, "A2", "placeholder")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:customSum(expr="total" lastCell="A2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "custom_cmd.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	data := map[string]any{"total": 42}

	out, err := FillBytes(tmpl, data, WithCommand("customSum", func(attrs map[string]string) (Command, error) {
		return &customSumCommand{Expr: attrs["expr"]}, nil
	}))
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	v, _ := outFile.GetCellValue(sheet, "A2")
	assert.Equal(t, "42", v)
}
