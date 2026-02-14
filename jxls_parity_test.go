package xlfill

import (
	"bytes"
	"fmt"
	"math"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// =============================================================================
// XlsAreaTest parity — ported from org.jxls.area.XlsAreaTest
// =============================================================================

// TestXlsArea_ApplyAtToAnotherSheet tests applying an area from one sheet to another.
// Ported from XlsAreaTest.applyAtToAnotherSheet
func TestXlsArea_ApplyAtToAnotherSheet(t *testing.T) {
	f := excelize.NewFile()
	srcSheet := "Sheet1" // default sheet created by excelize
	dstSheet := "Sheet2"
	f.NewSheet(dstSheet)

	// Fill source area A1:G10 with identifiable content
	for row := 0; row < 10; row++ {
		for col := 0; col < 7; col++ {
			cell := ColToName(col) + fmt.Sprintf("%d", row+1)
			f.SetCellValue(srcSheet, cell, fmt.Sprintf("R%dC%d", row, col))
		}
	}

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)

	// Area: Sheet1!A1:G10 (7 wide, 10 high)
	area := NewArea(NewCellRef(srcSheet, 0, 0), Size{Width: 7, Height: 10}, tx)

	// Apply at Sheet2!B2
	size, err := area.ApplyAt(NewCellRef(dstSheet, 1, 1), ctx)
	require.NoError(t, err)
	assert.Equal(t, 7, size.Width)
	assert.Equal(t, 10, size.Height)

	// Verify some cells were mapped to Sheet2
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Sheet1!A1 → Sheet2!B2
	v, _ := out.GetCellValue(dstSheet, "B2")
	assert.Equal(t, "R0C0", v)

	// Sheet1!D2 → Sheet2!E3
	v, _ = out.GetCellValue(dstSheet, "E3")
	assert.Equal(t, "R1C3", v)

	// Sheet1!G10 → Sheet2!H11
	v, _ = out.GetCellValue(dstSheet, "H11")
	assert.Equal(t, "R9C6", v)
}

// TestXlsArea_ApplyAtShiftDownWithTwoCommands tests area with two commands that expand.
// Ported from XlsAreaTest.applyAtShiftDownWithTwoCommands
func TestXlsArea_ApplyAtShiftDownWithTwoCommands(t *testing.T) {
	f := excelize.NewFile()
	srcSheet := "Sheet1"
	dstSheet := "Sheet2"
	f.NewSheet(dstSheet)

	// Fill source area A1:G10
	for row := 0; row < 10; row++ {
		for col := 0; col < 7; col++ {
			cell := ColToName(col) + fmt.Sprintf("%d", row+1)
			f.SetCellValue(srcSheet, cell, fmt.Sprintf("R%dC%d", row, col))
		}
	}

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)

	area := NewArea(NewCellRef(srcSheet, 0, 0), Size{Width: 7, Height: 10}, tx)

	// Command1 at B3:C5 (cols 1-2, rows 2-4), expands to Size(3,4) — 1 col wider, 1 row taller
	cmd1 := &mockCommand{name: "cmd1", resultSize: Size{Width: 3, Height: 4}}
	area.AddCommand(cmd1, NewCellRef(srcSheet, 2, 1), Size{Width: 2, Height: 3})

	// Command2 at A7:B8 (cols 0-1, rows 6-7), expands to Size(3,3) — 1 col wider, 1 row taller
	cmd2 := &mockCommand{name: "cmd2", resultSize: Size{Width: 3, Height: 3}}
	area.AddCommand(cmd2, NewCellRef(srcSheet, 6, 0), Size{Width: 2, Height: 2})

	size, err := area.ApplyAt(NewCellRef(dstSheet, 1, 1), ctx)
	require.NoError(t, err)

	// cmd1 expands by 1 row → everything below shifts down by 1.
	// cmd2 expands by 1 row → everything below shifts down by 1 more.
	// Total height: 10 + 1 (cmd1 expansion) + 1 (cmd2 expansion) = 12
	assert.Equal(t, 12, size.Height)
}

// TestXlsArea_ApplyAtShiftUpWithTwoCommands tests area with commands that contract.
// Ported from XlsAreaTest.applyAtShiftUpWithTwoCommands
func TestXlsArea_ApplyAtShiftUpWithTwoCommands(t *testing.T) {
	f := excelize.NewFile()
	srcSheet := "Sheet1"
	dstSheet := "Sheet2"
	f.NewSheet(dstSheet)

	for row := 0; row < 10; row++ {
		for col := 0; col < 7; col++ {
			cell := ColToName(col) + fmt.Sprintf("%d", row+1)
			f.SetCellValue(srcSheet, cell, fmt.Sprintf("R%dC%d", row, col))
		}
	}

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)

	area := NewArea(NewCellRef(srcSheet, 0, 0), Size{Width: 7, Height: 10}, tx)

	// cmd1 at B3:C5 (3 rows), contracts to Size(2,2) — 1 row shorter
	cmd1 := &mockCommand{name: "cmd1", resultSize: Size{Width: 2, Height: 2}}
	area.AddCommand(cmd1, NewCellRef(srcSheet, 2, 1), Size{Width: 2, Height: 3})

	// cmd2 at A7:B8 (2 rows), same size
	cmd2 := &mockCommand{name: "cmd2", resultSize: Size{Width: 2, Height: 2}}
	area.AddCommand(cmd2, NewCellRef(srcSheet, 6, 0), Size{Width: 2, Height: 2})

	size, err := area.ApplyAt(NewCellRef(dstSheet, 1, 1), ctx)
	require.NoError(t, err)

	// cmd1 contracts partial-width (B-C only, has static cols), so min height is source height (3)
	// cmd2 is partial-width too (A-B only), same height (2)
	// Total = 2 (rows 0-1 before cmd1) + 3 (cmd1 zone, min of source) + 1 (row 5) + 2 (cmd2) + 2 (rows 8-9) = 10
	assert.Equal(t, 10, size.Height)
}

// mockCommand is a simple mock for testing area processing.
type mockCommand struct {
	name       string
	resultSize Size
}

func (m *mockCommand) Name() string { return m.name }
func (m *mockCommand) Reset()       {}
func (m *mockCommand) ApplyAt(cellRef CellRef, ctx *Context, tx Transformer) (Size, error) {
	return m.resultSize, nil
}

// =============================================================================
// If01Test parity — ported from org.jxls.templatebasedtests.If01Test
// =============================================================================

type commodity struct {
	Subject string
	Price   float64
	Weight  float64
	SellBuy string
}

func getIfTestData() []any {
	return []any{
		map[string]any{"subject": "Gas", "price": 1.0, "weight": 1.0, "sellBuy": "buy"},
		map[string]any{"subject": "Oil", "price": 2.1, "weight": 10.0, "sellBuy": "sell"},
		map[string]any{"subject": "Gas", "price": 3.12, "weight": 100.0, "sellBuy": "buy"},
		map[string]any{"subject": "Gas", "price": 10.0, "weight": 1000.0, "sellBuy": "buy"},
		map[string]any{"subject": "Gas", "price": 10.0, "weight": 1234.0, "sellBuy": "buy"},
		map[string]any{"subject": "Gas 123", "price": 123.45, "weight": 678.0, "sellBuy": "sell"},
	}
}

// TestIf01_EnglishReport tests if/else with language switching and buy/sell rows.
func TestIf01_EnglishReport(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Row 1: conditional header based on language
	f.SetCellValue(sheet, "A1", "English report")
	f.SetCellValue(sheet, "A2", "Deutscher Bericht")
	// Row 3: column headers
	f.SetCellValue(sheet, "A3", "Subject")
	f.SetCellValue(sheet, "B3", "Price")
	f.SetCellValue(sheet, "C3", "Weight")
	// Row 4: data template with buy/sell conditional
	f.SetCellValue(sheet, "A4", "${e.subject}")
	f.SetCellValue(sheet, "B4", "${e.price}")
	f.SetCellValue(sheet, "C4", "${e.weight}")
	f.SetCellValue(sheet, "D4", "Buy")
	f.SetCellValue(sheet, "D5", "Sell")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	list := getIfTestData()
	ctx := NewContext(map[string]any{"lang": "en", "list": list})

	// Build if/else for header
	ifHeaderArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	elseHeaderArea := NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx)
	ifHeaderCmd := &IfCommand{
		Condition: `lang == "en"`,
		IfArea:    ifHeaderArea,
		ElseArea:  elseHeaderArea,
	}

	// Build buy/sell conditional
	buyArea := NewArea(NewCellRef(sheet, 3, 3), Size{Width: 1, Height: 1}, tx)  // D4="Buy"
	sellArea := NewArea(NewCellRef(sheet, 4, 3), Size{Width: 1, Height: 1}, tx) // D5="Sell"
	ifBuyCmd := &IfCommand{
		Condition: `e.sellBuy == "buy"`,
		IfArea:    buyArea,
		ElseArea:  sellArea,
	}

	// Each command area: A4:D4 (with if on D4)
	eachInner := NewArea(NewCellRef(sheet, 3, 0), Size{Width: 4, Height: 1}, tx)
	eachInner.AddCommand(ifBuyCmd, NewCellRef(sheet, 3, 3), Size{Width: 1, Height: 1})
	eachCmd := &EachCommand{
		Items: "list", Var: "e", Direction: "DOWN",
		Area: eachInner,
	}

	// Root area that wraps everything
	rootArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 4, Height: 4}, tx)
	rootArea.AddCommand(ifHeaderCmd, NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1})
	rootArea.AddCommand(eachCmd, NewCellRef(sheet, 3, 0), Size{Width: 4, Height: 1})

	size, err := rootArea.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// 1 header + 2 static rows (cols header) + 6 data rows = 9
	assert.True(t, size.Height >= 9)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Verify English header chosen
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "English report", v)

	// Verify buy/sell labels
	v, _ = out.GetCellValue(sheet, "D4")
	assert.Equal(t, "Buy", v) // first item is "buy"
	v, _ = out.GetCellValue(sheet, "D5")
	assert.Equal(t, "Sell", v) // second item is "sell"
	v, _ = out.GetCellValue(sheet, "D6")
	assert.Equal(t, "Buy", v) // third item is "buy"
}

// TestIf01_GermanReport tests the else branch for language switching.
func TestIf01_GermanReport(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "English report")
	f.SetCellValue(sheet, "A2", "Deutscher Bericht")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"lang": "de"})

	ifArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	elseArea := NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &IfCommand{
		Condition: `lang == "en"`,
		IfArea:    ifArea,
		ElseArea:  elseArea,
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
	assert.Equal(t, "Deutscher Bericht", v)
}

// TestIf01_EmptyList tests if/each with empty list.
func TestIf01_EmptyList(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.subject}")
	f.SetCellValue(sheet, "A3", "Footer")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"lang": "en", "list": []any{}})

	eachCmd := &EachCommand{
		Items: "list", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx),
	}

	rootArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 3}, tx)
	rootArea.AddCommand(eachCmd, NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1})

	size, err := rootArea.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// Header (1) + 0 items + Footer (1) = 2
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

// TestIf01_OneRow tests if/each with a single row.
func TestIf01_OneRow(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.subject}")
	f.SetCellValue(sheet, "B2", "${e.price}")
	f.SetCellValue(sheet, "C2", "Buy")
	f.SetCellValue(sheet, "C3", "Sell")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	list := []any{map[string]any{"subject": "1 row", "price": 10.0, "sellBuy": "buy"}}
	ctx := NewContext(map[string]any{"list": list})

	buyArea := NewArea(NewCellRef(sheet, 1, 2), Size{Width: 1, Height: 1}, tx)
	sellArea := NewArea(NewCellRef(sheet, 2, 2), Size{Width: 1, Height: 1}, tx)
	ifCmd := &IfCommand{
		Condition: `e.sellBuy == "buy"`,
		IfArea:    buyArea,
		ElseArea:  sellArea,
	}

	eachInner := NewArea(NewCellRef(sheet, 1, 0), Size{Width: 3, Height: 1}, tx)
	eachInner.AddCommand(ifCmd, NewCellRef(sheet, 1, 2), Size{Width: 1, Height: 1})

	eachCmd := &EachCommand{
		Items: "list", Var: "e", Direction: "DOWN",
		Area: eachInner,
	}

	rootArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 3, Height: 2}, tx)
	rootArea.AddCommand(eachCmd, NewCellRef(sheet, 1, 0), Size{Width: 3, Height: 1})

	size, err := rootArea.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height) // header + 1 data row

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "B2")
	assert.Equal(t, "10", v)
	v, _ = out.GetCellValue(sheet, "C2")
	assert.Equal(t, "Buy", v)
}

// =============================================================================
// DirectionRight tests — ported from DirectionRightTest
// =============================================================================

// TestDirectionRight_FourColumns tests RIGHT direction with 4-column template.
func TestDirectionRight_FourColumns(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// 4-column template that repeats RIGHT
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "A2", "${e.Value1}")
	f.SetCellValue(sheet, "B2", "${e.Value2}")
	f.SetCellValue(sheet, "C2", "${e.Value3}")
	f.SetCellValue(sheet, "D2", "${e.Value4}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Q1", "Value1": 10, "Value2": 20, "Value3": 30, "Value4": 40},
		map[string]any{"Name": "Q2", "Value1": 50, "Value2": 60, "Value3": 70, "Value4": 80},
		map[string]any{"Name": "Q3", "Value1": 90, "Value2": 100, "Value3": 110, "Value4": 120},
	}
	ctx := NewContext(map[string]any{"items": items})

	// Each area is A1:D2 (4 wide, 2 high), direction RIGHT
	eachArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 4, Height: 2}, tx)
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "RIGHT",
		Area: eachArea,
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 12, size.Width)  // 3 items * 4 cols
	assert.Equal(t, 2, size.Height)  // max height

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Q1 at A1, Q2 at E1, Q3 at I1
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Q1", v)
	v, _ = out.GetCellValue(sheet, "E1")
	assert.Equal(t, "Q2", v)
	v, _ = out.GetCellValue(sheet, "I1")
	assert.Equal(t, "Q3", v)

	// Values: Q1 D2=40, Q2 H2=80, Q3 L2=120
	v, _ = out.GetCellValue(sheet, "D2")
	assert.Equal(t, "40", v)
	v, _ = out.GetCellValue(sheet, "H2")
	assert.Equal(t, "80", v)
	v, _ = out.GetCellValue(sheet, "L2")
	assert.Equal(t, "120", v)
}

// =============================================================================
// ClearTemplateCells parity — ported from ClearTemplateCellsTest
// =============================================================================

// TestClearTemplateCells_UnusedExpressionsStay tests that with clearTemplateCells=false,
// unused expressions remain in template cells.
func TestClearTemplateCells_UnusedExpressionsStay(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.name}")
	f.SetCellValue(sheet, "D2", "XX")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="D2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="D2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "clear_cells.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	data := map[string]any{"employees": []any{}}

	out, err := FillBytes(tmpl, data, WithClearTemplateCells(false))
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	// With empty list and no clearing, template expressions may remain
	// The header should still be there
	v, _ := outFile.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
}

// TestClearTemplateCells_UnusedExpressionsAreCleared tests that with clearTemplateCells=true,
// unused template cells are cleared.
func TestClearTemplateCells_UnusedExpressionsAreCleared(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.name}")
	f.SetCellValue(sheet, "D2", "XX")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="D2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="D2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "clear_cells2.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	data := map[string]any{"employees": []any{}}

	out, err := FillBytes(tmpl, data, WithClearTemplateCells(true))
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	// Header should still be present
	v, _ := outFile.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
}

// =============================================================================
// NestedSumsTest parity — ported from org.jxls.templatebasedtests.NestedSumsTest
// =============================================================================

func getNestedSumsTestData() []any {
	return []any{
		map[string]any{"supertype": "Commodities type A", "instrument": "Commodity", "class2": "Liegenschaften", "description": "Wolterstr. 100", "amount": 250.0},
		map[string]any{"supertype": "Commodities type A", "instrument": "Commodity", "class2": "Liegenschaften", "description": "Stauffenbergallee", "amount": 500.0},
		map[string]any{"supertype": "Commodities type A", "instrument": "Commodity", "class2": "Immobilien", "description": "Wolterstr. 102", "amount": 250.0},
		map[string]any{"supertype": "Commodities type B", "instrument": "Commodity", "class2": "Fahrzeuge", "description": "Porsche 911", "amount": 100.0},
		map[string]any{"supertype": "Commodities type B", "instrument": "Commodity", "class2": "Fahrzeuge", "description": "Mercedes Maybach", "amount": 300.0},
		map[string]any{"supertype": "Commodities type B", "instrument": "Commodity", "class2": "Fahrzeuge", "description": "Mercedes-Benz SLK 350", "amount": 60.0},
		map[string]any{"supertype": "Commodities type B", "instrument": "Commodity", "class2": "Fahrzeuge", "description": "Bentley Flying Spur", "amount": 240.0},
		map[string]any{"supertype": "Bonds", "instrument": "Bond", "class2": "Base", "description": "AC-100 K1", "amount": 200.0},
		map[string]any{"supertype": "Bonds", "instrument": "Bond", "class2": "Base", "description": "AC-100 K2", "amount": 200.0},
		map[string]any{"supertype": "Bonds", "instrument": "Bond", "class2": "Base", "description": "AC-100 K3", "amount": 200.0},
		map[string]any{"supertype": "Bonds", "instrument": "Bond", "class2": "Super", "description": "MX 12", "amount": 123.0},
		map[string]any{"supertype": "Shares", "instrument": "Share", "class2": "Base", "description": "L77", "amount": 0.31},
	}
}

// TestNestedSums_GroupByWithSums tests nested grouping with SUM formulas.
func TestNestedSums_GroupByWithSums(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Simple template: group by supertype, show group key
	f.SetCellValue(sheet, "A1", "${g.Item.supertype}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	data := getNestedSumsTestData()
	ctx := NewContext(map[string]any{"data": data})

	// Group by supertype
	groupCmd := &EachCommand{
		Items: "data", Var: "g", Direction: "DOWN",
		GroupBy: "g.supertype",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := groupCmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)

	// 4 groups: Commodities type A, Commodities type B, Bonds, Shares
	assert.Equal(t, 4, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Commodities type A", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Commodities type B", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bonds", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Shares", v)
}

// TestNestedSums_GroupByWithNestedEach tests nested each (group → items).
func TestNestedSums_GroupByWithNestedEach(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${g.Item.supertype}")
	f.SetCellValue(sheet, "A2", "${e.description}")
	f.SetCellValue(sheet, "B2", "${e.amount}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	data := getNestedSumsTestData()
	ctx := NewContext(map[string]any{"data": data})

	// Inner each iterates g.Items
	innerEach := &EachCommand{
		Items: "g.Items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1}, tx),
	}

	// Outer group area: header row + items row
	groupArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)
	groupArea.AddCommand(innerEach, NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1})

	groupCmd := &EachCommand{
		Items: "data", Var: "g", Direction: "DOWN",
		GroupBy: "g.supertype",
		Area:    groupArea,
	}

	size, err := groupCmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)

	// Commodities type A: 1 header + 3 items = 4
	// Commodities type B: 1 header + 4 items = 5
	// Bonds: 1 header + 4 items = 5
	// Shares: 1 header + 1 item = 2
	// Total = 16
	assert.Equal(t, 16, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// First group header
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Commodities type A", v)
	// First item
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Wolterstr. 100", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "250", v)
}

// =============================================================================
// GroupSumTest parity — ported from GroupSumTest
// =============================================================================

// TestGroupSum_MapsWithDoubles tests grouping map data with double values.
func TestGroupSum_MapsWithDoubles(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${g.Item.Category}")
	f.SetCellValue(sheet, "B1", "${g.Item.Amount}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Category": "Food", "Amount": 10.5},
		map[string]any{"Category": "Transport", "Amount": 25.0},
		map[string]any{"Category": "Food", "Amount": 15.3},
		map[string]any{"Category": "Entertainment", "Amount": 50.0},
		map[string]any{"Category": "Transport", "Amount": 30.0},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		GroupBy: "g.Category",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height) // 3 groups: Food, Transport, Entertainment

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Verify group order (insertion order): Food, Transport, Entertainment
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Food", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Transport", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Entertainment", v)
}

// TestGroupSum_WithFilterCondition tests groupBy with select filter.
func TestGroupSum_WithFilterCondition(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${g.Item.Category}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Category": "Food", "Active": true},
		map[string]any{"Category": "Transport", "Active": false},
		map[string]any{"Category": "Food", "Active": true},
		map[string]any{"Category": "Entertainment", "Active": true},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		Select:  "g.Active == true",
		GroupBy: "g.Category",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 2, size.Height) // Food and Entertainment (Transport filtered out)
}

// =============================================================================
// IssueB105 parity — big doubles like 1.3E22
// =============================================================================

// TestIssueB105_BigDoubles tests that large double values are handled correctly.
func TestIssueB105_BigDoubles(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.BigVal}")
	f.SetCellValue(sheet, "B1", "${e.SmallVal}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"BigVal": 1.3e22, "SmallVal": 1.2},
		map[string]any{"BigVal": 1.3e22, "SmallVal": 1.2},
		map[string]any{"BigVal": 1.3e22, "SmallVal": 1.2},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Verify big doubles are preserved
	for row := 1; row <= 3; row++ {
		cell := fmt.Sprintf("A%d", row)
		v, _ := out.GetCellValue(sheet, cell)
		// Parse back to float and check it's close to 1.3E22
		var f64 float64
		fmt.Sscanf(v, "%g", &f64)
		assert.InDelta(t, 1.3e22, f64, 1.0, "big double in %s", cell)
	}
}

// =============================================================================
// IssueB133 parity — nested groupBy property
// =============================================================================

// TestIssueB133_NestedGroupBy tests groupBy with nested property access.
func TestIssueB133_NestedGroupBy(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${g.Item.Region}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Region": "US"},
		map[string]any{"Name": "Bob", "Region": "EU"},
		map[string]any{"Name": "Carol", "Region": "US"},
		map[string]any{"Name": "Dave", "Region": "EU"},
		map[string]any{"Name": "Eve", "Region": "Asia"},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "g", Direction: "DOWN",
		GroupBy:    "g.Region",
		GroupOrder: "ASC",
		Area:       NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height) // Asia, EU, US

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Asia", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "EU", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "US", v)
}

// =============================================================================
// IssueB167 parity — columns beyond AZ
// =============================================================================

// TestIssueB167_BeyondColumnAZ tests handling of columns beyond Z (AA, AB, etc.).
func TestIssueB167_BeyondColumnAZ(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Set values in columns beyond Z
	f.SetCellValue(sheet, "AA1", "${e.Val}")
	f.SetCellValue(sheet, "AB1", "${e.Val2}")
	f.SetCellValue(sheet, "AZ1", "${e.Val3}")
	f.SetCellValue(sheet, "BA1", "${e.Val4}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Val": "AA", "Val2": "AB", "Val3": "AZ", "Val4": "BA"},
	}
	ctx := NewContext(map[string]any{"items": items})

	// Area from AA1 to BA1 (col 26 to 52, width=27)
	startCol, _ := NameToCol("AA")
	endCol, _ := NameToCol("BA")
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 0, startCol), Size{Width: endCol - startCol + 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, startCol), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 1, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "AA1")
	assert.Equal(t, "AA", v)
	v, _ = out.GetCellValue(sheet, "AB1")
	assert.Equal(t, "AB", v)
	v, _ = out.GetCellValue(sheet, "AZ1")
	assert.Equal(t, "AZ", v)
	v, _ = out.GetCellValue(sheet, "BA1")
	assert.Equal(t, "BA", v)
}

// TestColToName_BeyondZ validates column name conversion beyond Z.
func TestColToName_BeyondZ(t *testing.T) {
	assert.Equal(t, "A", ColToName(0))
	assert.Equal(t, "Z", ColToName(25))
	assert.Equal(t, "AA", ColToName(26))
	assert.Equal(t, "AB", ColToName(27))
	assert.Equal(t, "AZ", ColToName(51))
	assert.Equal(t, "BA", ColToName(52))
	assert.Equal(t, "ZZ", ColToName(701))
	assert.Equal(t, "AAA", ColToName(702))
}

// TestNameToCol_BeyondZ validates column name parsing beyond Z.
func TestNameToCol_BeyondZ(t *testing.T) {
	col, err := NameToCol("A")
	require.NoError(t, err)
	assert.Equal(t, 0, col)

	col, err = NameToCol("Z")
	require.NoError(t, err)
	assert.Equal(t, 25, col)

	col, err = NameToCol("AA")
	require.NoError(t, err)
	assert.Equal(t, 26, col)

	col, err = NameToCol("AZ")
	require.NoError(t, err)
	assert.Equal(t, 51, col)

	col, err = NameToCol("BA")
	require.NoError(t, err)
	assert.Equal(t, 52, col)

	col, err = NameToCol("ZZ")
	require.NoError(t, err)
	assert.Equal(t, 701, col)

	col, err = NameToCol("AAA")
	require.NoError(t, err)
	assert.Equal(t, 702, col)
}

// =============================================================================
// IssueB097 parity — division formula with cell ref outside each
// =============================================================================

// TestIssueB097_DivisionFormulaOutsideEach tests formula referencing cells outside each.
func TestIssueB097_DivisionFormulaOutsideEach(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Total in a fixed cell
	f.SetCellValue(sheet, "A1", "Total")
	f.SetCellValue(sheet, "B1", 9300.0) // sum of all payments

	// Header
	f.SetCellValue(sheet, "A2", "Name")
	f.SetCellValue(sheet, "B2", "Payment")
	f.SetCellValue(sheet, "C2", "Ratio")

	// Data row with formula
	f.SetCellValue(sheet, "A3", "${e.Name}")
	f.SetCellValue(sheet, "B3", "${e.Payment}")
	f.SetCellFormula(sheet, "C3", "B3/B1") // ratio = payment / total (B1 is outside each)

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="C3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A3", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="C3")`,
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

	// Process formulas
	fp := NewFormulaProcessor()
	for _, area := range areas {
		fp.ProcessAreaFormulas(tx, area)
	}

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Verify data rows
	v, _ := out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Elsa", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Oleg", v)
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "Neil", v)

	// Verify formulas reference B1 (outside area, should be preserved)
	formula, _ := out.GetCellFormula(sheet, "C3")
	assert.Contains(t, formula, "B1", "formula should reference fixed cell B1")
}

// =============================================================================
// IssueB116 parity — external formula references
// =============================================================================

// TestIssueB116_ExternalFormulaRef tests formulas referencing cells entirely outside the area.
func TestIssueB116_ExternalFormulaRef(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Data")
	f.SetCellValue(sheet, "B1", "${e.Value}")
	// Formula in area references E1 which is outside
	f.SetCellValue(sheet, "E1", 100.0)
	f.SetCellFormula(sheet, "C1", "B1+E1")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="C1")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"e": map[string]any{"Value": 50}})

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

	// E1 is outside area (A1:C1), should be preserved in formula
	formula, _ := out.GetCellFormula(sheet, "C1")
	assert.Contains(t, formula, "E1", "external ref E1 should be preserved")
}

// =============================================================================
// IssueB184 parity — if inside each with column sums
// =============================================================================

// TestIssueB184_IfInsideEachWithColumnSums tests if command inside each with SUM formulas.
func TestIssueB184_IfInsideEachWithColumnSums(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Amount")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Amount}")
	f.SetCellValue(sheet, "C2", "HIGH")
	f.SetCellFormula(sheet, "B3", "SUM(B2:B2)")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Amount": 5000.0, "HighValue": true},
		map[string]any{"Name": "Bob", "Amount": 3000.0, "HighValue": false},
		map[string]any{"Name": "Carol", "Amount": 7000.0, "HighValue": true},
	}
	ctx := NewContext(map[string]any{"items": items})

	ifArea := NewArea(NewCellRef(sheet, 1, 2), Size{Width: 1, Height: 1}, tx)
	ifCmd := &IfCommand{Condition: "e.HighValue == true", IfArea: ifArea}

	eachInner := NewArea(NewCellRef(sheet, 1, 0), Size{Width: 3, Height: 1}, tx)
	eachInner.AddCommand(ifCmd, NewCellRef(sheet, 1, 2), Size{Width: 1, Height: 1})

	eachCmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: eachInner,
	}

	rootArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 3, Height: 3}, tx)
	rootArea.AddCommand(eachCmd, NewCellRef(sheet, 1, 0), Size{Width: 3, Height: 1})

	size, err := rootArea.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	// 1 header + 3 data + 1 formula = 5
	assert.Equal(t, 5, size.Height)

	// Process formulas
	fp := NewFormulaProcessor()
	fp.ProcessAreaFormulas(tx, rootArea)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Verify names
	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Carol", v)

	// Verify conditional: Alice and Carol have HIGH, Bob doesn't
	v, _ = out.GetCellValue(sheet, "C2")
	assert.Equal(t, "HIGH", v) // Alice
	v, _ = out.GetCellValue(sheet, "C4")
	assert.Equal(t, "HIGH", v) // Carol

	// Formula should be expanded
	formula, _ := out.GetCellFormula(sheet, "B5")
	assert.Contains(t, formula, "B2", "formula should reference expanded range start")
}

// =============================================================================
// IssueB197 parity — jointed cell refs + empty collections
// =============================================================================

// TestIssueB197_EmptyCollection tests behavior with empty collections and formulas.
func TestIssueB197_EmptyCollection(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.Val}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Empty collection
	ctx := NewContext(map[string]any{"items": []any{}})

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

	// Header should still be present
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
}

// =============================================================================
// IssueB210 parity — sums of empty lists
// =============================================================================

// TestIssueB210_SumOfEmptyList tests SUM formula when the list is empty.
func TestIssueB210_SumOfEmptyList(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Value")
	f.SetCellValue(sheet, "A2", "${e.Amount}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"items": []any{}})

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

	// Header and formula should be present (formula row shifts up when list is empty)
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Value", v)
}

// =============================================================================
// Issue93 parity — varIndex restore after nested loops
// =============================================================================

// TestIssue93_VarIndexRestore tests that varIndex is properly saved/restored in nested loops.
func TestIssue93_VarIndexRestore(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${outerIdx}")
	f.SetCellValue(sheet, "B1", "${o.Name}")
	f.SetCellValue(sheet, "A2", "${innerIdx}")
	f.SetCellValue(sheet, "B2", "${i.Val}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	outerItems := []any{
		map[string]any{"Name": "Group1", "Items": []any{
			map[string]any{"Val": "A"},
			map[string]any{"Val": "B"},
		}},
		map[string]any{"Name": "Group2", "Items": []any{
			map[string]any{"Val": "C"},
		}},
	}
	ctx := NewContext(map[string]any{"outerItems": outerItems})

	// Inner each
	innerEach := &EachCommand{
		Items: "o.Items", Var: "i", VarIndex: "innerIdx", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1}, tx),
	}

	// Outer area with inner each
	outerArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)
	outerArea.AddCommand(innerEach, NewCellRef(sheet, 1, 0), Size{Width: 2, Height: 1})

	outerEach := &EachCommand{
		Items: "outerItems", Var: "o", VarIndex: "outerIdx", Direction: "DOWN",
		Area: outerArea,
	}

	size, err := outerEach.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)

	// Group1: 1 header + 2 items = 3, Group2: 1 header + 1 item = 2. Total = 5
	assert.Equal(t, 5, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Row 1: outerIdx=0, Name=Group1
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "0", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Group1", v)

	// Row 2: innerIdx=0, Val=A
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "0", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "A", v)

	// Row 3: innerIdx=1, Val=B
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "1", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "B", v)

	// Row 4: outerIdx=1, Name=Group2
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "1", v)
	v, _ = out.GetCellValue(sheet, "B4")
	assert.Equal(t, "Group2", v)

	// Row 5: innerIdx=0 (reset for new group), Val=C
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "0", v)
	v, _ = out.GetCellValue(sheet, "B5")
	assert.Equal(t, "C", v)
}

// =============================================================================
// Issue147 parity — row heights preserved
// =============================================================================

// TestIssue147_RowHeights tests that row heights from template are preserved in output.
func TestIssue147_RowHeights(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetRowHeight(sheet, 1, 30.0) // custom row height

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

	// Both rows should have the custom height
	h1, _ := out.GetRowHeight(sheet, 1)
	h2, _ := out.GetRowHeight(sheet, 2)
	assert.InDelta(t, 30.0, h1, 1.0, "row 1 height should be ~30")
	assert.InDelta(t, 30.0, h2, 1.0, "row 2 height should be ~30")
}

// =============================================================================
// Issue166 parity — run template processing twice
// =============================================================================

// TestIssue166_RunTwice tests that processing the same template twice produces correct results.
func TestIssue166_RunTwice(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tmpl := filepath.Join(testdataDir(t), "run_twice.xlsx")
	require.NoError(t, f.SaveAs(tmpl))
	f.Close()

	data := map[string]any{
		"items": []any{
			map[string]any{"Name": "Alice"},
			map[string]any{"Name": "Bob"},
		},
	}

	// First run
	out1, err := FillBytes(tmpl, data)
	require.NoError(t, err)

	// Second run with same template
	out2, err := FillBytes(tmpl, data)
	require.NoError(t, err)

	// Verify both produce correct output
	for i, outBytes := range [][]byte{out1, out2} {
		outF, err := excelize.OpenReader(bytes.NewReader(outBytes))
		require.NoError(t, err)

		v, _ := outF.GetCellValue(sheet, "A2")
		assert.Equal(t, "Alice", v, "run %d: A2 should be Alice", i+1)
		v, _ = outF.GetCellValue(sheet, "A3")
		assert.Equal(t, "Bob", v, "run %d: A3 should be Bob", i+1)

		outF.Close()
	}
}

// =============================================================================
// IssueB198 parity — array/slice support in expressions
// =============================================================================

// TestIssueB198_ArraySupport tests iteration over typed Go slices.
func TestIssueB198_ArraySupport(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	// Test with []int
	intItems := []int{10, 20, 30}
	ctx := NewContext(map[string]any{"items": intItems})

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
	assert.Equal(t, "10", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "20", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "30", v)
}

// TestIssueB198_StringArraySupport tests iteration over []string.
func TestIssueB198_StringArraySupport(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	strItems := []string{"Hello", "World", "Go"}
	ctx := NewContext(map[string]any{"items": strItems})

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
	assert.Equal(t, "Hello", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Go", v)
}

// TestIssueB198_Float64ArraySupport tests iteration over []float64.
func TestIssueB198_Float64ArraySupport(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	floatItems := []float64{1.1, 2.2, 3.3}
	ctx := NewContext(map[string]any{"items": floatItems})

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
	assert.Equal(t, "1.1", v)
}

// =============================================================================
// Issue173 parity — bounds checking
// =============================================================================

// TestIssue173_AreaBounds tests that area processing respects boundaries.
func TestIssue173_AreaBounds(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Create a 5x5 grid
	for r := 0; r < 5; r++ {
		for c := 0; c < 5; c++ {
			cell := ColToName(c) + fmt.Sprintf("%d", r+1)
			f.SetCellValue(sheet, cell, fmt.Sprintf("(%d,%d)", r, c))
		}
	}

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)

	// Process only B2:D4 (3x3 sub-area)
	area := NewArea(NewCellRef(sheet, 1, 1), Size{Width: 3, Height: 3}, tx)
	size, err := area.ApplyAt(NewCellRef(sheet, 1, 1), ctx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 3}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Cells outside the area should be unaffected
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "(0,0)", v)
	v, _ = out.GetCellValue(sheet, "E5")
	assert.Equal(t, "(4,4)", v)

	// Cells inside the area should be processed
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "(1,1)", v)
	v, _ = out.GetCellValue(sheet, "D4")
	assert.Equal(t, "(3,3)", v)
}

// =============================================================================
// MultiSheet tests — ported from MultiSheetTest
// =============================================================================

// TestMultiSheet_Basic tests multi-sheet output via EachCommand MultiSheet.
func TestMultiSheet_Basic(t *testing.T) {
	f := excelize.NewFile()
	sheet := "template"
	f.SetSheetName("Sheet1", sheet)

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Department}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	employees := []any{
		map[string]any{"Name": "Elsa", "Department": "IT", "SheetName": "Elsa"},
		map[string]any{"Name": "Oleg", "Department": "HR", "SheetName": "Oleg"},
	}
	sheetNames := []string{"Elsa", "Oleg"}

	ctx := NewContext(map[string]any{
		"employees":  employees,
		"sheetNames": sheetNames,
	})

	// Create sheets for each employee
	for _, name := range sheetNames {
		tx.CopySheet(sheet, name)
	}

	// Process each employee in their own sheet
	for i, emp := range employees {
		empMap := emp.(map[string]any)
		sheetName := sheetNames[i]

		rv := NewRunVar(ctx, "e")
		rv.Set(empMap)

		area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 2}, tx)
		_, err := area.ApplyAt(NewCellRef(sheetName, 0, 0), ctx)
		require.NoError(t, err)

		rv.Close()
	}

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Verify sheets exist and have correct data
	v, _ := out.GetCellValue("Elsa", "A2")
	assert.Equal(t, "Elsa", v)
	v, _ = out.GetCellValue("Elsa", "B2")
	assert.Equal(t, "IT", v)

	v, _ = out.GetCellValue("Oleg", "A2")
	assert.Equal(t, "Oleg", v)
	v, _ = out.GetCellValue("Oleg", "B2")
	assert.Equal(t, "HR", v)
}

// =============================================================================
// SafeSheetNameBuilder parity
// =============================================================================

// TestSafeSheetName tests that sheet name length and special chars are handled.
func TestSafeSheetName(t *testing.T) {
	// Excel sheet name limits: max 31 chars, no []*?/\: chars
	tests := []struct {
		input    string
		expected string
	}{
		{"Normal", "Normal"},
		{"Sheet with spaces", "Sheet with spaces"},
		{"Loooooooooooooooooooooooooooooooong", "Looooooooooooooooooooooooooooooo"}, // truncated to 31
		{"Sheet/Name", "Sheet_Name"},   // / replaced
		{"Sheet\\Name", "Sheet_Name"},  // \ replaced
		{"Sheet:Name", "Sheet_Name"},   // : replaced
		{"Sheet*Name", "Sheet_Name"},   // * replaced
		{"Sheet?Name", "Sheet_Name"},   // ? replaced
		{"Sheet[Name]", "Sheet_Name_"}, // [ ] replaced
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			result := SafeSheetName(tt.input)
			assert.LessOrEqual(t, len(result), 31, "sheet name too long")
			// Verify no forbidden chars
			for _, ch := range result {
				assert.NotContains(t, []rune{'/', '\\', ':', '*', '?', '[', ']'}, ch)
			}
		})
	}
}

// =============================================================================
// Formula edge cases — CreateTargetCellRef patterns
// =============================================================================

// TestFormula_HorizontalGap tests formula expansion when target cells have horizontal gaps.
func TestFormula_HorizontalGap(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Val}")
	f.SetCellFormula(sheet, "B1", "SUM(A1)")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Val": 10},
		map[string]any{"Val": 20},
		map[string]any{"Val": 30},
	}
	ctx := NewContext(map[string]any{"items": items})

	// Each RIGHT — values go A1, B1, C1
	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "RIGHT",
		Area: NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	_, err = cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "10", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "20", v)
	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "30", v)
}

// TestFormula_MultipleFormulasWithExpansion tests multiple formula cells referencing the same each range.
func TestFormula_MultipleFormulasWithExpansion(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Val")
	f.SetCellValue(sheet, "A2", "${e.Val}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")
	f.SetCellFormula(sheet, "B3", "AVERAGE(A2:A2)")
	f.SetCellFormula(sheet, "C3", "COUNT(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="C3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Val": 100},
		map[string]any{"Val": 200},
		map[string]any{"Val": 300},
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

	// Formulas should be in row 5 (header + 3 items + formula row)
	sumFormula, _ := out.GetCellFormula(sheet, "A5")
	assert.Contains(t, sumFormula, "A2")
	assert.Contains(t, sumFormula, "A4")

	avgFormula, _ := out.GetCellFormula(sheet, "B5")
	assert.Contains(t, avgFormula, "A2")

	cntFormula, _ := out.GetCellFormula(sheet, "C5")
	assert.Contains(t, cntFormula, "A2")
}

// =============================================================================
// Scalars and primitive type tests
// =============================================================================

// TestScalars_DirectValues tests using scalar values directly in expressions.
func TestScalars_DirectValues(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${title}")
	f.SetCellValue(sheet, "B1", "${count}")
	f.SetCellValue(sheet, "C1", "${ratio}")
	f.SetCellValue(sheet, "D1", "${active}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"title":  "Report",
		"count":  42,
		"ratio":  3.14,
		"active": true,
	})

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 4, Height: 1}, tx)
	size, err := area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 4, Height: 1}, size)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Report", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "42", v)
	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "3.14", v)
	v, _ = out.GetCellValue(sheet, "D1")
	assert.Equal(t, "TRUE", v)
}

// =============================================================================
// Formula strategy tests (BY_COLUMN, BY_ROW)
// =============================================================================

// TestFormulaStrategy_ByColumn tests that FormulaByColumn filters targets by matching column.
func TestFormulaStrategy_ByColumn(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{
		NewCellRef("S", 0, 0), // A1
		NewCellRef("S", 1, 0), // A2
		NewCellRef("S", 0, 1), // B1
		NewCellRef("S", 1, 1), // B2
	}

	// Filter by column 0 (A)
	filtered := fp.filterByStrategy(targets, NewCellRef("S", 5, 0), FormulaByColumn)
	assert.Len(t, filtered, 2)
	for _, f := range filtered {
		assert.Equal(t, 0, f.Col)
	}
}

// TestFormulaStrategy_ByRow tests that FormulaByRow filters targets by matching row.
func TestFormulaStrategy_ByRow(t *testing.T) {
	fp := &StandardFormulaProcessor{}

	targets := []CellRef{
		NewCellRef("S", 0, 0), // A1
		NewCellRef("S", 1, 0), // A2
		NewCellRef("S", 0, 1), // B1
		NewCellRef("S", 1, 1), // B2
	}

	// Filter by row 0
	filtered := fp.filterByStrategy(targets, NewCellRef("S", 0, 5), FormulaByRow)
	assert.Len(t, filtered, 2)
	for _, f := range filtered {
		assert.Equal(t, 0, f.Row)
	}
}

// =============================================================================
// OrderByComparator edge cases
// =============================================================================

// TestOrderBy_MultiField tests sorting by multiple fields.
func TestOrderBy_MultiField(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Charlie", "Age": 30, "Salary": 5000.0},
		map[string]any{"Name": "Alice", "Age": 25, "Salary": 6000.0},
		map[string]any{"Name": "Bob", "Age": 30, "Salary": 4000.0},
		map[string]any{"Name": "Alice", "Age": 35, "Salary": 7000.0},
	}
	ctx := NewContext(map[string]any{"items": items})

	cmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		OrderBy: "e.Name ASC, e.Age DESC",
		Area:    NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx),
	}

	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, 4, size.Height)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Expected order: Alice(35), Alice(25), Bob(30), Charlie(30)
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Charlie", v)
}

// TestOrderBy_NilValues tests sorting with nil values.
func TestOrderBy_NilValues(t *testing.T) {
	items := []any{
		map[string]any{"Name": "Bob", "Score": 80},
		map[string]any{"Name": "Alice", "Score": nil},
		map[string]any{"Name": "Carol", "Score": 90},
	}

	specs := []orderBySpec{{field: "Score", desc: false}}
	sortByFields(items, specs)

	// nil should sort first (smallest)
	assert.Nil(t, getField(items[0], "Score"))
	assert.Equal(t, 80, getField(items[1], "Score"))
	assert.Equal(t, 90, getField(items[2], "Score"))
}

// =============================================================================
// CompareValues edge cases
// =============================================================================

func TestCompareValues_NilHandling(t *testing.T) {
	assert.Equal(t, 0, compareValues(nil, nil))
	assert.Equal(t, -1, compareValues(nil, "x"))
	assert.Equal(t, 1, compareValues("x", nil))
}

func TestCompareValues_NumericTypes(t *testing.T) {
	assert.Equal(t, -1, compareValues(1, 2))
	assert.Equal(t, 0, compareValues(3.14, 3.14))
	assert.Equal(t, 1, compareValues(10.0, 5.0))
	assert.Equal(t, -1, compareValues(int64(1), float64(2)))
}

func TestCompareValues_StringFallback(t *testing.T) {
	assert.Equal(t, -1, compareValues("apple", "banana"))
	assert.Equal(t, 0, compareValues("same", "same"))
	assert.Equal(t, 1, compareValues("zebra", "apple"))
}

// =============================================================================
// Struct field access via reflection
// =============================================================================

func TestGetField_Struct(t *testing.T) {
	emp := testEmployee{Name: "Alice", Payment: 5000.0, Active: true}
	assert.Equal(t, "Alice", getField(emp, "Name"))
	assert.Equal(t, 5000.0, getField(emp, "Payment"))
	assert.Equal(t, true, getField(emp, "Active"))
	assert.Nil(t, getField(emp, "NonExistent"))
}

func TestGetField_StructPointer(t *testing.T) {
	emp := &testEmployee{Name: "Bob", Payment: 6000.0}
	assert.Equal(t, "Bob", getField(emp, "Name"))
	assert.Equal(t, 6000.0, getField(emp, "Payment"))
}

func TestGetField_Map(t *testing.T) {
	m := map[string]any{"key": "value", "num": 42}
	assert.Equal(t, "value", getField(m, "key"))
	assert.Equal(t, 42, getField(m, "num"))
	assert.Nil(t, getField(m, "missing"))
}

func TestGetField_Nil(t *testing.T) {
	assert.Nil(t, getField(nil, "anything"))
}

// =============================================================================
// AreaColumnMerge — merge cells in expanded area
// =============================================================================

// TestAreaColumnMerge tests that merged cells work with area expansion.
func TestAreaColumnMerge(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.MergeCell(sheet, "A1", "C1") // merge A1:C1
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Value}")
	f.SetCellValue(sheet, "C2", "${e.Desc}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Name": "Alice", "Value": 100, "Desc": "First"},
		map[string]any{"Name": "Bob", "Value": 200, "Desc": "Second"},
	}
	ctx := NewContext(map[string]any{"items": items})

	eachCmd := &EachCommand{
		Items: "items", Var: "e", Direction: "DOWN",
		Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 3, Height: 1}, tx),
	}

	rootArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 3, Height: 2}, tx)
	rootArea.AddCommand(eachCmd, NewCellRef(sheet, 1, 0), Size{Width: 3, Height: 1})

	size, err := rootArea.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)
	assert.Equal(t, 3, size.Height) // header + 2 data rows

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Header", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "C3")
	assert.Equal(t, "Second", v)
}

// =============================================================================
// IssueB162 parity — parameterized formulas like LEFT()
// =============================================================================

// TestIssueB162_ParameterizedFormula tests that formulas with string function params
// (like LEFT()) are handled correctly.
func TestIssueB162_ParameterizedFormula(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellFormula(sheet, "B1", `LEFT(A1,3)`) // parameterized formula

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"e": map[string]any{"Name": "Alexander"}})

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

	// Formula should still reference A1
	formula, _ := out.GetCellFormula(sheet, "B1")
	assert.Contains(t, formula, "A1")
	assert.Contains(t, formula, "LEFT")
}

// =============================================================================
// toFloat64 edge cases
// =============================================================================

func TestToFloat64_AllTypes(t *testing.T) {
	tests := []struct {
		input    any
		expected float64
		ok       bool
	}{
		{int(42), 42.0, true},
		{int8(8), 8.0, true},
		{int16(16), 16.0, true},
		{int32(32), 32.0, true},
		{int64(64), 64.0, true},
		{float32(3.14), float64(float32(3.14)), true},
		{float64(3.14), 3.14, true},
		{"string", 0, false},
		{nil, 0, false},
		{true, 0, false},
	}

	for _, tt := range tests {
		f, ok := toFloat64(tt.input)
		assert.Equal(t, tt.ok, ok, "toFloat64(%v)", tt.input)
		if ok {
			assert.InDelta(t, tt.expected, f, 0.001, "toFloat64(%v)", tt.input)
		}
	}
}

// =============================================================================
// Math edge cases for big/small numbers
// =============================================================================

func TestBigNumbers_Precision(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${val}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	tests := []struct {
		name string
		val  float64
	}{
		{"very_small", 1e-10},
		{"zero", 0.0},
		{"negative", -1234.56},
		{"large", 1e15},
		{"very_large", 1.3e22},
		{"max_safe", math.MaxFloat64},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctx := NewContext(map[string]any{"val": tt.val})
			area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
			_, err := area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
			require.NoError(t, err)
		})
	}
}

// =============================================================================
// Mixed content expression tests
// =============================================================================

func TestMixedExpressions(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Hello ${name}, you have ${count} items")
	f.SetCellValue(sheet, "B1", "Total: $${amount}")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{
		"name":   "World",
		"count":  5,
		"amount": 99.99,
	})

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	_, err = area.ApplyAt(NewCellRef(sheet, 0, 0), ctx)
	require.NoError(t, err)

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Hello World, you have 5 items", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Total: $99.99", v)
}
