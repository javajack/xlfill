package goxls

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// createIntegrationTemplate creates a full integration test template.
// Layout:
//
//	A1: "Name"           B1: "Age"       C1: "Salary"     [comment: jx:area(lastCell="C2")]
//	A2: "${e.Name}"      B2: "${e.Age}"  C2: "${e.Salary}" [comment: jx:each(items="employees" var="e" lastCell="C2")]
func createIntegrationTemplate(t *testing.T) string {
	t.Helper()
	return createBasicTemplate(t)
}

// createIfIntegrationTemplate creates a template with each + if.
// Layout:
//
//	A1: "Name"         B1: "Payment"     C1: "Status"       [comment: jx:area(lastCell="C2")]
//	A2: "${e.Name}"    B2: "${e.Payment}" C2: "HIGH"         [comment: jx:each(items="employees" var="e" lastCell="C2")]
//	                                      C2 also has: jx:if(condition="e.Payment > 2000" lastCell="C2")
func createIfIntegrationTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Payment")
	f.SetCellValue(sheet, "C1", "Status")

	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Payment}")
	f.SetCellValue(sheet, "C2", "HIGH")

	f.AddComment(sheet, excelize.Comment{
		Cell:   "A1",
		Author: "goxls",
		Text:   `jx:area(lastCell="C2")`,
	})
	// Each on A2, if on C2 — both in same row
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A2",
		Author: "goxls",
		Text:   `jx:each(items="employees" var="e" lastCell="C2")`,
	})

	path := filepath.Join(testdataDir(t), "if_integration.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func TestFill_BasicEach(t *testing.T) {
	tmpl := createIntegrationTemplate(t)
	outPath := filepath.Join(testdataDir(t), "out_basic_each.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000.0},
			map[string]any{"Name": "Bob", "Age": 25, "Salary": 6000.0},
			map[string]any{"Name": "Carol", "Age": 35, "Salary": 7000.0},
		},
	}

	err := Fill(tmpl, outPath, data)
	require.NoError(t, err)

	// Read output
	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	// Header row
	v, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "Name", v)
	v, _ = f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "Age", v)

	// Data rows
	v, _ = f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v)
	v, _ = f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "30", v)
	v, _ = f.GetCellValue("Sheet1", "C2")
	assert.Equal(t, "5000", v)

	v, _ = f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Bob", v)
	v, _ = f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "Carol", v)
}

func TestFill_EmptyList(t *testing.T) {
	tmpl := createIntegrationTemplate(t)
	outPath := filepath.Join(testdataDir(t), "out_empty.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{"employees": []any{}}

	err := Fill(tmpl, outPath, data)
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	// Header should still be present
	v, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "Name", v)
}

func TestFill_OutputToWriter(t *testing.T) {
	tmpl := createIntegrationTemplate(t)

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000.0},
		},
	}

	var buf bytes.Buffer
	filler := NewFiller(WithTemplate(tmpl))
	err := filler.FillWriter(data, &buf)
	require.NoError(t, err)
	assert.True(t, buf.Len() > 0)

	// Verify content
	f, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v)
}

func TestFill_OutputToBytes(t *testing.T) {
	tmpl := createIntegrationTemplate(t)

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Bob", "Age": 40, "Salary": 8000.0},
		},
	}

	out, err := FillBytes(tmpl, data)
	require.NoError(t, err)
	assert.True(t, len(out) > 0)

	f, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Bob", v)
}

func TestFill_PreservesFormatting(t *testing.T) {
	tmpl := createIntegrationTemplate(t)

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000.0},
			map[string]any{"Name": "Bob", "Age": 25, "Salary": 6000.0},
		},
	}

	out, err := FillBytes(tmpl, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer f.Close()

	// Header row should have bold style (from createBasicTemplate)
	s, _ := f.GetCellStyle("Sheet1", "A1")
	assert.True(t, s > 0, "header should have bold style")
}

func TestFill_InvalidTemplate(t *testing.T) {
	err := Fill("/nonexistent/template.xlsx", "/tmp/out.xlsx", map[string]any{})
	assert.Error(t, err)
}

func TestFill_TemplateFromReader(t *testing.T) {
	tmpl := createIntegrationTemplate(t)

	templateFile, err := os.Open(tmpl)
	require.NoError(t, err)
	defer templateFile.Close()

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "ReaderTest", "Age": 99, "Salary": 1.0},
		},
	}

	var buf bytes.Buffer
	err = FillReader(templateFile, &buf, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(&buf)
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "ReaderTest", v)
}

func TestFill_MapData(t *testing.T) {
	tmpl := createIntegrationTemplate(t)

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "MapUser", "Age": 50, "Salary": 9999.0},
		},
	}

	out, err := FillBytes(tmpl, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "MapUser", v)
	v, _ = f.GetCellValue("Sheet1", "C2")
	assert.Equal(t, "9999", v)
}

func TestFill_StructData(t *testing.T) {
	type Employee struct {
		Name   string
		Age    int
		Salary float64
	}

	tmpl := createIntegrationTemplate(t)

	data := map[string]any{
		"employees": []any{
			Employee{Name: "StructUser", Age: 28, Salary: 4500.0},
		},
	}

	out, err := FillBytes(tmpl, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "StructUser", v)
}

func TestFill_CustomNotation(t *testing.T) {
	// Create template with {{ }} notation
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "{{e.Name}}")

	f.AddComment(sheet, excelize.Comment{
		Cell:   "A1",
		Author: "goxls",
		Text:   `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A2",
		Author: "goxls",
		Text:   `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tmplPath := filepath.Join(testdataDir(t), "custom_notation.xlsx")
	require.NoError(t, f.SaveAs(tmplPath))
	f.Close()

	data := map[string]any{
		"items": []any{
			map[string]any{"Name": "CustomNotation"},
		},
	}

	out, err := FillBytes(tmplPath, data, WithExpressionNotation("{{", "}}"))
	require.NoError(t, err)

	outFile, err := excelize.OpenReader(bytes.NewReader(out))
	require.NoError(t, err)
	defer outFile.Close()

	v, _ := outFile.GetCellValue(sheet, "A2")
	assert.Equal(t, "CustomNotation", v)
}

func TestFill_PreWriteCallback(t *testing.T) {
	tmpl := createIntegrationTemplate(t)
	callbackCalled := false

	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Test", "Age": 1, "Salary": 1.0},
		},
	}

	filler := NewFiller(
		WithTemplate(tmpl),
		WithPreWrite(func(tx Transformer) error {
			callbackCalled = true
			return nil
		}),
	)

	_, err := filler.FillBytes(data)
	require.NoError(t, err)
	assert.True(t, callbackCalled)
}

func TestBuildAreas_SingleArea(t *testing.T) {
	tmpl := createIntegrationTemplate(t)
	tx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	areas, err := filler.BuildAreas(tx)
	require.NoError(t, err)
	require.Len(t, areas, 1)

	area := areas[0]
	assert.Equal(t, "Sheet1", area.StartCell.Sheet)
	assert.Equal(t, 0, area.StartCell.Row)
	assert.Equal(t, 0, area.StartCell.Col)
	assert.Equal(t, Size{Width: 3, Height: 2}, area.AreaSize)

	// Should have one command binding (the jx:each)
	require.Len(t, area.Bindings, 1)
	assert.Equal(t, "each", area.Bindings[0].Command.Name())
}

func TestBuildAreas_NoAreaCommand(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "No area here")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	_, err = filler.BuildAreas(tx)
	assert.Error(t, err)
}

func TestBuildAreas_NoComments(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Plain cell")
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	filler := NewFiller()
	_, err = filler.BuildAreas(tx)
	assert.Error(t, err)
}
