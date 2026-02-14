package goxls

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// createBenchTemplate creates a benchmark template file.
func createBenchTemplate(b *testing.B) string {
	b.Helper()
	f := excelize.NewFile()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "ID")
	f.SetCellValue(sheet, "B1", "Name")
	f.SetCellValue(sheet, "C1", "Value")

	f.SetCellValue(sheet, "A2", "${e.ID}")
	f.SetCellValue(sheet, "B2", "${e.Name}")
	f.SetCellValue(sheet, "C2", "${e.Value}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="C2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="items" var="e" lastCell="C2")`,
	})

	dir := filepath.Join("testdata")
	path := filepath.Join(dir, "bench_template.xlsx")
	if err := f.SaveAs(path); err != nil {
		b.Fatal(err)
	}
	f.Close()
	return path
}

func benchFill(b *testing.B, numRows int) {
	tmpl := createBenchTemplate(b)
	items := make([]any, numRows)
	for i := range items {
		items[i] = map[string]any{
			"ID":    i + 1,
			"Name":  fmt.Sprintf("Employee_%d", i),
			"Value": float64(i) * 1.5,
		}
	}
	data := map[string]any{"items": items}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := FillBytes(tmpl, data)
		if err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkFill_100Rows(b *testing.B)   { benchFill(b, 100) }
func BenchmarkFill_1000Rows(b *testing.B)  { benchFill(b, 1000) }
func BenchmarkFill_10000Rows(b *testing.B) { benchFill(b, 10000) }

func BenchmarkFill_NestedLoops(b *testing.B) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${d.Name}")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	path := filepath.Join("testdata", "bench_nested.xlsx")
	require.NoError(b, f.SaveAs(path))
	f.Close()

	departments := make([]any, 10)
	for i := range departments {
		emps := make([]any, 20)
		for j := range emps {
			emps[j] = map[string]any{"Name": fmt.Sprintf("Emp_%d_%d", i, j)}
		}
		departments[i] = map[string]any{
			"Name":      fmt.Sprintf("Dept_%d", i),
			"Employees": emps,
		}
	}

	// Manually build template since nested each requires wiring
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ef, _ := excelize.OpenFile(path)
		tx, _ := NewExcelizeTransformer(ef)
		ctx := NewContext(map[string]any{"departments": departments})

		innerEach := &EachCommand{
			Items: "d.Employees", Var: "e", Direction: "DOWN",
			Area: NewArea(NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1}, tx),
		}
		outerArea := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 2}, tx)
		outerArea.AddCommand(innerEach, NewCellRef(sheet, 1, 0), Size{Width: 1, Height: 1})
		outerEach := &EachCommand{
			Items: "departments", Var: "d", Direction: "DOWN",
			Area: outerArea,
		}

		outerEach.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
		tx.Close()
	}
}

func BenchmarkExprEvaluate(b *testing.B) {
	eval := NewExpressionEvaluator()
	data := map[string]any{
		"e": map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000.0},
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		eval.Evaluate("e.Name", data)
	}
}

func BenchmarkParseComment(b *testing.B) {
	comment := `jx:each(items="employees" var="e" lastCell="C2")`
	ref := NewCellRef("Sheet1", 0, 0)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ParseComment(comment, ref)
	}
}
