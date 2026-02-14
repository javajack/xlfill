package goxls

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestFormulaProcessor_SimpleSum(t *testing.T) {
	// Template: A1=header, A2=${e.Amount} (each), A3=SUM(A2:A2)
	// After 4 items: A2-A5 have data, A6 should have SUM(A2:A5)
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Amount")
	f.SetCellValue(sheet, "A2", "${e.Amount}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"Amount": 100.0},
		map[string]any{"Amount": 200.0},
		map[string]any{"Amount": 300.0},
		map[string]any{"Amount": 400.0},
	}
	ctx := NewContext(map[string]any{"items": items})

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)

	for _, area := range areas {
		_, err := area.ApplyAt(area.StartCell, ctx)
		require.NoError(t, err)
	}

	// Now process formulas
	fp := NewFormulaProcessor()
	for _, area := range areas {
		fp.ProcessAreaFormulas(tx, area)
	}

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Data should be in A2:A5
	v, _ := out.GetCellValue(sheet, "A2")
	assert.Equal(t, "100", v)
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "400", v)

	// Formula in A6 should reference expanded range
	formula, _ := out.GetCellFormula(sheet, "A6")
	// Should be SUM(A2:A5) or similar expanded form
	assert.Contains(t, formula, "A2")
	assert.Contains(t, formula, "A5")
}

func TestFormulaProcessor_ExternalRef(t *testing.T) {
	// Formula referencing cells outside the area should be preserved as-is.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Data")
	f.SetCellValue(sheet, "B1", 100)
	f.SetCellFormula(sheet, "A2", "B1*2") // B1 is outside jx:area A1:A2

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(nil)
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

	formula, _ := out.GetCellFormula(sheet, "A2")
	assert.Contains(t, formula, "B1")
}

func TestFormulaProcessor_NoFormulas(t *testing.T) {
	// Template with no formulas — processor should not crash.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{map[string]any{"Name": "Alice"}}
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
		fp.ProcessAreaFormulas(tx, area) // should not panic
	}
}

func TestFormulaProcessor_SingleCellRef(t *testing.T) {
	// A formula like "=A2+1" where A2 maps to a single target.
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${e.Val}")
	f.SetCellFormula(sheet, "B1", "A1+1")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="B1")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"e": map[string]any{"Val": 10}})

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

	formula, _ := out.GetCellFormula(sheet, "B1")
	assert.Contains(t, formula, "A1")
}

func TestParseCellRefFromFormula(t *testing.T) {
	ref, err := parseCellRefFromFormula("A1", "Sheet1")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ref.Sheet)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 0, ref.Col)

	ref, err = parseCellRefFromFormula("$B$5", "Sheet1")
	require.NoError(t, err)
	assert.Equal(t, 4, ref.Row)
	assert.Equal(t, 1, ref.Col)

	ref, err = parseCellRefFromFormula("Sheet2!C3", "Sheet1")
	require.NoError(t, err)
	assert.Equal(t, "Sheet2", ref.Sheet)
	assert.Equal(t, 2, ref.Row)
	assert.Equal(t, 2, ref.Col)
}

func TestCellRefRegex(t *testing.T) {
	matches := cellRefRegex.FindAllString("SUM(A1:B5)", -1)
	assert.Len(t, matches, 2) // A1, B5

	matches = cellRefRegex.FindAllString("Sheet1!C3+$D$4", -1)
	assert.Len(t, matches, 2)

	matches = cellRefRegex.FindAllString("A1+B2*C3", -1)
	assert.Len(t, matches, 3)
}

func TestFormulaProcessor_VerticalRange(t *testing.T) {
	// SUM(A2:A2) with 3-item expansion → SUM(A2:A4)
	fp := NewFormulaProcessor()

	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Val")
	f.SetCellValue(sheet, "A2", "${e.V}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	items := []any{
		map[string]any{"V": 10},
		map[string]any{"V": 20},
		map[string]any{"V": 30},
	}
	ctx := NewContext(map[string]any{"items": items})

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)

	for _, area := range areas {
		_, err := area.ApplyAt(area.StartCell, ctx)
		require.NoError(t, err)
		fp.ProcessAreaFormulas(tx, area)
	}

	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	out, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer out.Close()

	// Formula should be in A5 (header + 3 data rows + formula)
	formula, _ := out.GetCellFormula(sheet, "A5")
	assert.Contains(t, formula, "A2")
	assert.Contains(t, formula, "A4")
}
