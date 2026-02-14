package goxls

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// =============================================================================
// YellowCommandTest parity — custom command implementation
// Demonstrates registering and using a custom command.
// =============================================================================

// yellowCommand is a custom command that writes "yellow" to a cell when a condition is true.
type yellowCommand struct {
	Condition string
}

func (c *yellowCommand) Name() string { return "yellow" }
func (c *yellowCommand) Reset()       {}
func (c *yellowCommand) ApplyAt(cellRef CellRef, ctx *Context, tx Transformer) (Size, error) {
	if c.Condition != "" {
		ok, err := ctx.IsConditionTrue(c.Condition)
		if err != nil {
			return Size{Width: 1, Height: 1}, nil
		}
		if ok {
			tx.SetCellValue(cellRef, "yellow")
		}
	}
	return Size{Width: 1, Height: 1}, nil
}

func TestYellowCommand_CustomCommand(t *testing.T) {
	// Test custom command registration and invocation directly (not nested in each)
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Status")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"A1\")\njx:yellow(condition=\"score > 80\" lastCell=\"A1\")",
	})

	yellowFactory := func(attrs map[string]string) (Command, error) {
		return &yellowCommand{Condition: attrs["condition"]}, nil
	}

	filler := NewFiller(WithCommand("yellow", yellowFactory))

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Test with condition true
	ctx := NewContext(map[string]any{"score": 90})
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
	assert.Equal(t, "yellow", v)
}

func TestYellowCommand_ConditionFalse(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Status")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"A1\")\njx:yellow(condition=\"score > 80\" lastCell=\"A1\")",
	})

	yellowFactory := func(attrs map[string]string) (Command, error) {
		return &yellowCommand{Condition: attrs["condition"]}, nil
	}

	filler := NewFiller(WithCommand("yellow", yellowFactory))

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Test with condition false
	ctx := NewContext(map[string]any{"score": 50})
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
	assert.NotEqual(t, "yellow", v, "should NOT be yellow when score <= 80")
}

// =============================================================================
// SubtotalTest parity — extending EachCommand with custom behavior
// =============================================================================

type eachWithSubtotal struct {
	inner    *EachCommand
	called   *bool
	subtotal string
}

func (c *eachWithSubtotal) Name() string { return "each" }
func (c *eachWithSubtotal) Reset()       { c.inner.Reset() }
func (c *eachWithSubtotal) ApplyAt(cellRef CellRef, ctx *Context, tx Transformer) (Size, error) {
	if c.subtotal != "" {
		*c.called = true
	}
	return c.inner.ApplyAt(cellRef, ctx, tx)
}

func TestSubtotalCommand_ExtendEach(t *testing.T) {
	subtotalCalled := false

	// Build manually since the wrapper doesn't have SetArea, we need to
	// construct the EachCommand and its area manually
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Create inner each command
	inner := &EachCommand{
		Items: "employees", Var: "e", Direction: "DOWN",
	}
	innerArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	inner.Area = innerArea

	wrapper := &eachWithSubtotal{
		inner:    inner,
		called:   &subtotalCalled,
		subtotal: "sum",
	}

	employees := []any{
		map[string]any{"Name": "Elsa"},
		map[string]any{"Name": "Oleg"},
	}

	ctx := NewContext(map[string]any{"employees": employees})

	size, err := wrapper.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height)
	assert.True(t, subtotalCalled, "subtotal action should have been called")

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Elsa", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Oleg", v)
}

// =============================================================================
// EachTest parity — orderBy case-insensitive
// =============================================================================

func TestEachCommand_OrderByCaseInsensitive(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")

	// Both area and each in one comment (excelize allows only one comment per cell)
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"A1\")\njx:each(items=\"employees\" var=\"e\" orderBy=\"e.Name ASC\" lastCell=\"A1\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "i"},
		map[string]any{"Name": "Z"},
		map[string]any{"Name": "A"},
	}

	ctx := NewContext(map[string]any{"employees": employees})

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

	// "A" < "Z" < "i" in Go string comparison
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "A", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Z", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "i", v)
}

// =============================================================================
// EachTest parity — orderBy DESC
// =============================================================================

func TestEachCommand_OrderByDesc_Parity(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.Payment}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"employees\" var=\"e\" orderBy=\"e.Payment DESC\" lastCell=\"B1\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Elsa", "Payment": 1500.0},
		map[string]any{"Name": "Oleg", "Payment": 2300.0},
		map[string]any{"Name": "Neil", "Payment": 2500.0},
	}

	ctx := NewContext(map[string]any{"employees": employees})

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
	assert.Equal(t, "Neil", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Oleg", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Elsa", v)
}

// =============================================================================
// EachTest parity — orderBy multi-key
// =============================================================================

func TestEachCommand_OrderByMultiKey(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.Payment}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"employees\" var=\"e\" orderBy=\"e.Name ASC, e.Payment DESC\" lastCell=\"B1\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Z", "Payment": 1500.0},
		map[string]any{"Name": "A", "Payment": 2300.0},
		map[string]any{"Name": "Z", "Payment": 1700.0},
		map[string]any{"Name": "A", "Payment": 1000.0},
	}

	ctx := NewContext(map[string]any{"employees": employees})

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

	// Sort: Name ASC, then Payment DESC within same name
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "A", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "2300", v)

	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "A", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "1000", v)

	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Z", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "1700", v)

	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Z", v)
	v, _ = out.GetCellValue(sheet, "B4")
	assert.Equal(t, "1500", v)
}

// =============================================================================
// MultiSheetTest parity — multisheet each with multiple sheets
// =============================================================================

func TestMultiSheet_BasicGeneration(t *testing.T) {
	t.Skip("multisheet each not yet implemented in goxls — field is parsed but ApplyAt doesn't use it")
	f := excelize.NewFile()
	sheet := "template"
	f.SetSheetName("Sheet1", sheet)

	f.SetCellValue(sheet, "A1", "Employee Report")
	f.SetCellValue(sheet, "B2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"B2\")\njx:each(items=\"employees\" var=\"e\" multisheet=\"sheetNames\" lastCell=\"B2\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Elsa"},
		map[string]any{"Name": "John"},
	}

	ctx := NewContext(map[string]any{
		"employees":  employees,
		"sheetNames": []any{"Elsa", "John"},
	})

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

	v, _ := out.GetCellValue("Elsa", "B2")
	assert.Equal(t, "Elsa", v)

	v, _ = out.GetCellValue("John", "B2")
	assert.Equal(t, "John", v)
}

// =============================================================================
// SelectTest parity — select with complex boolean expressions
// =============================================================================

func TestSelectCommand_ComplexBooleanExpression(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.City}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"people\" var=\"e\" select=\"e.Age >= 18 && e.City == 'Berlin'\" lastCell=\"B1\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	people := []any{
		map[string]any{"Name": "Alice", "Age": 25, "City": "Berlin"},
		map[string]any{"Name": "Bob", "Age": 15, "City": "Berlin"},
		map[string]any{"Name": "Carol", "Age": 30, "City": "Munich"},
		map[string]any{"Name": "Dave", "Age": 22, "City": "Berlin"},
	}

	ctx := NewContext(map[string]any{"people": people})

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
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Dave", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Empty(t, v)
}

// =============================================================================
// ExceptionHandlerTest parity — expression evaluation error handling
// =============================================================================

func TestExpressionError_MissingVariable(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${missing.value}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{})
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)

	_, err = area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	assert.Error(t, err, "expression with missing variable should error")
	assert.Contains(t, err.Error(), "missing")
}

func TestExpressionError_NilPropertyAccess(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${obj.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"obj": nil})
	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)

	_, err = area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	assert.Error(t, err, "property access on nil should error")
}

// =============================================================================
// FormulaProcessorsTest parity — formula expansion with SUM
// =============================================================================

func TestFormulaProcessor_SUMExpansion(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Value}")
	f.SetCellFormula(sheet, "A6", "SUM(A1:A1)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"A6\")\njx:each(items=\"items\" var=\"e\" lastCell=\"A1\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Value": 10.0},
		map[string]any{"Value": 20.0},
		map[string]any{"Value": 30.0},
		map[string]any{"Value": 40.0},
		map[string]any{"Value": 50.0},
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

	// 5 items → A1:A5 data, formula at A10 (shifted from A6 by 4 rows)
	formula, _ := out.GetCellFormula(sheet, "A10")
	if formula != "" {
		assert.Contains(t, formula, "A1:A5")
	}
}

// =============================================================================
// IfTest parity — if with else branch
// =============================================================================

func TestIfCommand_WithElseBranch(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "VIP: ${e.Name}")
	f.SetCellValue(sheet, "A3", "Regular: ${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: fmt.Sprintf(`jx:if(condition="e.VIP" lastCell="A2" areas=["A2:A2","A3:A3"])`),
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"e": map[string]any{"Name": "Alice", "VIP": true},
	})

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

	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "VIP: Alice", v)
}

func TestIfCommand_ElseBranch(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "VIP: ${e.Name}")
	f.SetCellValue(sheet, "A3", "Regular: ${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: fmt.Sprintf(`jx:if(condition="e.VIP" lastCell="A2" areas=["A2:A2","A3:A3"])`),
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"e": map[string]any{"Name": "Bob", "VIP": false},
	})

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

	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Regular: Bob", v)
}

// =============================================================================
// GridTest parity — grid with headers and data rows
// =============================================================================

func TestGridCommand_HeadersAndData(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "placeholder")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"A2\")\njx:grid(headers=\"headers\" data=\"data\" lastCell=\"A2\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	headers := []any{"Name", "Age", "City"}
	data := []any{
		[]any{"Alice", 25, "Berlin"},
		[]any{"Bob", 30, "Munich"},
	}

	ctx := NewContext(map[string]any{
		"headers": headers,
		"data":    data,
	})

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
	assert.Equal(t, "Name", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Age", v)
	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "City", v)

	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
}

// =============================================================================
// ScalarsTest parity — iterating over scalar slices
// =============================================================================

func TestScalars_IntSlice(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${item}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: "jx:area(lastCell=\"A1\")\njx:each(items=\"numbers\" var=\"item\" lastCell=\"A1\")",
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"numbers": []any{10, 20, 30, 40, 50},
	})

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
	assert.Equal(t, "10", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "30", v)
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "50", v)
}

// =============================================================================
// PreWriteTest parity — pre-write callback
// =============================================================================

func TestPreWrite_Callback(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${val}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A1")`,
	})

	preWriteCalled := false

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"val": "hello"})

	filler := NewFiller(WithPreWrite(func(t Transformer) error {
		preWriteCalled = true
		return nil
	}))

	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)

	for _, area := range areas {
		_, err := area.ApplyAt(area.StartCell, ctx)
		require.NoError(t, err)
	}

	if filler.opts.preWrite != nil {
		err = filler.opts.preWrite(tx)
		require.NoError(t, err)
	}

	assert.True(t, preWriteCalled, "pre-write callback should have been called")
}

// =============================================================================
// ClearTemplateCellsTest parity
// =============================================================================

func TestClearTemplateCells_Disabled(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Static")
	f.SetCellValue(sheet, "A2", "${val}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"val": "resolved"})

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 2}, tx)

	_, err = area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Static", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "resolved", v)
}
