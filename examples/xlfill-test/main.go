package main

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"

	"github.com/javajack/xlfill"
	"github.com/xuri/excelize/v2"
)

var inputDir = "input"
var outputDir = "output"

func main() {
	os.MkdirAll(inputDir, 0o755)
	os.MkdirAll(outputDir, 0o755)

	tests := []struct {
		name string
		fn   func(outDir string) error
	}{
		{"01_basic_each", testBasicEach},
		{"02_each_with_varindex", testEachVarIndex},
		{"03_each_direction_right", testEachDirectionRight},
		{"04_each_with_select", testEachSelect},
		{"05_each_with_orderby", testEachOrderBy},
		{"06_each_with_groupby", testEachGroupBy},
		{"07_if_command", testIfCommand},
		{"08_formulas", testFormulas},
		{"09_grid_command", testGridCommand},
		{"10_image_command", testImageCommand},
		{"11_merge_cells", testMergeCells},
		{"12_hyperlinks", testHyperlinks},
		{"13_nested_each", testNestedEach},
		{"14_multisheet", testMultiSheet},
		{"15_custom_notation", testCustomNotation},
		{"16_keep_template_sheet", testKeepTemplateSheet},
		{"17_autorowheight", testAutoRowHeight},
		{"18_fill_bytes", testFillBytes},
		{"19_fill_reader", testFillReader},
	}

	passed, failed := 0, 0
	for _, tt := range tests {
		fmt.Printf("%-35s ", tt.name)
		if err := tt.fn(outputDir); err != nil {
			fmt.Printf("FAIL: %v\n", err)
			failed++
		} else {
			fmt.Printf("OK\n")
			passed++
		}
	}

	fmt.Printf("\n%d passed, %d failed out of %d tests\n", passed, failed, len(tests))
	if failed > 0 {
		os.Exit(1)
	}
}

// helper: save template to input/ and fill to output/
func fillTemplate(f *excelize.File, tmplName, outPath string, data map[string]any, opts ...xlfill.Option) error {
	tmplPath := filepath.Join(inputDir, tmplName)
	if err := f.SaveAs(tmplPath); err != nil {
		return fmt.Errorf("save template: %w", err)
	}
	f.Close()
	return xlfill.Fill(tmplPath, outPath, data, opts...)
}

// helper: read cell from output
func readCell(path, sheet, cell string) (string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return "", err
	}
	defer f.Close()
	v, err := f.GetCellValue(sheet, cell)
	return v, err
}

// helper: expect cell value
func expectCell(path, sheet, cell, expected string) error {
	v, err := readCell(path, sheet, cell)
	if err != nil {
		return err
	}
	if v != expected {
		return fmt.Errorf("cell %s!%s: got %q, want %q", sheet, cell, v, expected)
	}
	return nil
}

// helper: create a small PNG image as []byte
func createTestPNG() []byte {
	img := image.NewRGBA(image.Rect(0, 0, 10, 10))
	for y := 0; y < 10; y++ {
		for x := 0; x < 10; x++ {
			img.Set(x, y, color.RGBA{R: 0, G: 100, B: 200, A: 255})
		}
	}
	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}

// ===== 01: Basic Each =====
func testBasicEach(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Name")
	f.SetCellValue(s, "B1", "Age")
	f.SetCellValue(s, "C1", "Salary")
	f.SetCellValue(s, "A2", "${e.Name}")
	f.SetCellValue(s, "B2", "${e.Age}")
	f.SetCellValue(s, "C2", "${e.Salary}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="C2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="employees" var="e" lastCell="C2")`})

	out := filepath.Join(outDir, "01_basic_each.xlsx")
	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000},
			map[string]any{"Name": "Bob", "Age": 25, "Salary": 6000},
			map[string]any{"Name": "Carol", "Age": 35, "Salary": 7000},
		},
	}
	if err := fillTemplate(f, "t01.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A2", "Alice"); err != nil {
		return err
	}
	if err := expectCell(out, s, "A4", "Carol"); err != nil {
		return err
	}
	return expectCell(out, s, "C3", "6000")
}

// ===== 02: Each with varIndex =====
func testEachVarIndex(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "#")
	f.SetCellValue(s, "B1", "Item")
	f.SetCellValue(s, "A2", "${idx + 1}")
	f.SetCellValue(s, "B2", "${e}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="B2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="items" var="e" varIndex="idx" lastCell="B2")`})

	out := filepath.Join(outDir, "02_varindex.xlsx")
	data := map[string]any{"items": []any{"Apple", "Banana", "Cherry"}}
	if err := fillTemplate(f, "t02.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A2", "1"); err != nil {
		return err
	}
	if err := expectCell(out, s, "B3", "Banana"); err != nil {
		return err
	}
	return expectCell(out, s, "A4", "3")
}

// ===== 03: Each Direction RIGHT =====
func testEachDirectionRight(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "${e}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"A1\")\njx:each(items=\"months\" var=\"e\" direction=\"RIGHT\" lastCell=\"A1\")"})

	out := filepath.Join(outDir, "03_direction_right.xlsx")
	data := map[string]any{"months": []any{"Jan", "Feb", "Mar", "Apr"}}
	if err := fillTemplate(f, "t03.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A1", "Jan"); err != nil {
		return err
	}
	if err := expectCell(out, s, "B1", "Feb"); err != nil {
		return err
	}
	return expectCell(out, s, "D1", "Apr")
}

// ===== 04: Each with select =====
func testEachSelect(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Name")
	f.SetCellValue(s, "B1", "Salary")
	f.SetCellValue(s, "A2", "${e.Name}")
	f.SetCellValue(s, "B2", "${e.Salary}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="B2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="employees" var="e" select="e.Salary >= 6000" lastCell="B2")`})

	out := filepath.Join(outDir, "04_select.xlsx")
	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Salary": 5000},
			map[string]any{"Name": "Bob", "Salary": 6000},
			map[string]any{"Name": "Carol", "Salary": 7000},
		},
	}
	if err := fillTemplate(f, "t04.xlsx", out, data); err != nil {
		return err
	}
	// Only Bob and Carol should appear
	if err := expectCell(out, s, "A2", "Bob"); err != nil {
		return err
	}
	return expectCell(out, s, "A3", "Carol")
}

// ===== 05: Each with orderBy =====
func testEachOrderBy(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Name")
	f.SetCellValue(s, "A2", "${e.Name}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="names" var="e" orderBy="e.Name DESC" lastCell="A2")`})

	out := filepath.Join(outDir, "05_orderby.xlsx")
	data := map[string]any{
		"names": []any{
			map[string]any{"Name": "Charlie"},
			map[string]any{"Name": "Alice"},
			map[string]any{"Name": "Bob"},
		},
	}
	if err := fillTemplate(f, "t05.xlsx", out, data); err != nil {
		return err
	}
	// DESC order: Charlie, Bob, Alice
	if err := expectCell(out, s, "A2", "Charlie"); err != nil {
		return err
	}
	if err := expectCell(out, s, "A3", "Bob"); err != nil {
		return err
	}
	return expectCell(out, s, "A4", "Alice")
}

// ===== 06: Each with groupBy =====
func testEachGroupBy(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "${g.Item.Department}")
	f.SetCellValue(s, "B1", "${g.Item.Name}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"employees\" var=\"g\" groupBy=\"g.Department\" lastCell=\"B1\")"})

	out := filepath.Join(outDir, "06_groupby.xlsx")
	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Department": "Engineering"},
			map[string]any{"Name": "Bob", "Department": "Sales"},
			map[string]any{"Name": "Carol", "Department": "Engineering"},
		},
	}
	if err := fillTemplate(f, "t06.xlsx", out, data); err != nil {
		return err
	}
	// Should have 2 groups: Engineering and Sales
	if err := expectCell(out, s, "A1", "Engineering"); err != nil {
		return err
	}
	return expectCell(out, s, "A2", "Sales")
}

// ===== 07: If Command =====
func testIfCommand(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Name")
	f.SetCellValue(s, "B1", "Status")
	f.SetCellValue(s, "A2", "${e.Name}")
	f.SetCellValue(s, "B2", "ACTIVE")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="B2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="employees" var="e" lastCell="B2")`})
	f.AddComment(s, excelize.Comment{Cell: "B2", Author: "xlfill", Text: `jx:if(condition="e.Active" lastCell="B2")`})

	out := filepath.Join(outDir, "07_if_command.xlsx")
	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Active": true},
			map[string]any{"Name": "Bob", "Active": false},
			map[string]any{"Name": "Carol", "Active": true},
		},
	}
	if err := fillTemplate(f, "t07.xlsx", out, data); err != nil {
		return err
	}
	// Alice active, Bob not, Carol active
	if err := expectCell(out, s, "B2", "ACTIVE"); err != nil {
		return err
	}
	return nil
}

// ===== 08: Formulas =====
func testFormulas(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Amount")
	f.SetCellValue(s, "A2", "${e.Amount}")
	f.SetCellFormula(s, "A3", "SUM(A2:A2)")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A3")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="items" var="e" lastCell="A2")`})

	out := filepath.Join(outDir, "08_formulas.xlsx")
	data := map[string]any{
		"items": []any{
			map[string]any{"Amount": 100},
			map[string]any{"Amount": 200},
			map[string]any{"Amount": 300},
		},
	}
	if err := fillTemplate(f, "t08.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A2", "100"); err != nil {
		return err
	}
	return expectCell(out, s, "A4", "300")
}

// ===== 09: Grid Command =====
func testGridCommand(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "placeholder")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"A2\")\njx:grid(headers=\"headers\" data=\"data\" lastCell=\"A2\")"})

	out := filepath.Join(outDir, "09_grid.xlsx")
	data := map[string]any{
		"headers": []any{"Name", "Age", "City"},
		"data": []any{
			[]any{"Alice", 30, "NYC"},
			[]any{"Bob", 25, "LA"},
		},
	}
	if err := fillTemplate(f, "t09.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A1", "Name"); err != nil {
		return err
	}
	if err := expectCell(out, s, "C1", "City"); err != nil {
		return err
	}
	return expectCell(out, s, "A2", "Alice")
}

// ===== 10: Image Command =====
func testImageCommand(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Logo below")
	f.SetCellValue(s, "A2", "")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:image(src="logo" imageType="PNG" lastCell="A2")`})

	out := filepath.Join(outDir, "10_image.xlsx")
	data := map[string]any{
		"logo": createTestPNG(),
	}
	if err := fillTemplate(f, "t10.xlsx", out, data); err != nil {
		return err
	}
	// Just verify the file was created and is valid
	of, err := excelize.OpenFile(out)
	if err != nil {
		return err
	}
	of.Close()
	return nil
}

// ===== 11: MergeCells =====
func testMergeCells(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Merged Header")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"C2\")\njx:mergeCells(lastCell=\"C2\" cols=\"3\" rows=\"2\")"})

	out := filepath.Join(outDir, "11_mergecells.xlsx")
	data := map[string]any{}
	if err := fillTemplate(f, "t11.xlsx", out, data); err != nil {
		return err
	}
	of, err := excelize.OpenFile(out)
	if err != nil {
		return err
	}
	defer of.Close()
	merges, _ := of.GetMergeCells(s)
	if len(merges) == 0 {
		return fmt.Errorf("expected merged cells, got none")
	}
	return nil
}

// ===== 12: Hyperlinks =====
func testHyperlinks(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Site")
	f.SetCellValue(s, "B1", "Link")
	f.SetCellValue(s, "A2", "${e.Name}")
	f.SetCellValue(s, "B2", "${hyperlink(e.URL, e.Name)}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="B2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="sites" var="e" lastCell="B2")`})

	out := filepath.Join(outDir, "12_hyperlinks.xlsx")
	data := map[string]any{
		"sites": []any{
			map[string]any{"Name": "Google", "URL": "https://google.com"},
			map[string]any{"Name": "GitHub", "URL": "https://github.com"},
		},
	}
	if err := fillTemplate(f, "t12.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A2", "Google"); err != nil {
		return err
	}
	return expectCell(out, s, "A3", "GitHub")
}

// ===== 13: Nested Each =====
func testNestedEach(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "${dept.Name}")
	f.SetCellValue(s, "A2", "${e.Name}")
	f.SetCellValue(s, "B2", "${e.Role}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"B2\")\njx:each(items=\"departments\" var=\"dept\" lastCell=\"B2\")"})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="dept.Employees" var="e" lastCell="B2")`})

	out := filepath.Join(outDir, "13_nested_each.xlsx")
	data := map[string]any{
		"departments": []any{
			map[string]any{
				"Name": "Engineering",
				"Employees": []any{
					map[string]any{"Name": "Alice", "Role": "Lead"},
					map[string]any{"Name": "Bob", "Role": "Dev"},
				},
			},
			map[string]any{
				"Name": "Sales",
				"Employees": []any{
					map[string]any{"Name": "Carol", "Role": "Manager"},
				},
			},
		},
	}
	if err := fillTemplate(f, "t13.xlsx", out, data); err != nil {
		return err
	}
	if err := expectCell(out, s, "A1", "Engineering"); err != nil {
		return err
	}
	return expectCell(out, s, "A2", "Alice")
}

// ===== 14: MultiSheet =====
func testMultiSheet(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "${dept.Name}")
	f.SetCellValue(s, "A2", "${dept.Head}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"A2\")\njx:each(items=\"departments\" var=\"dept\" multisheet=\"sheetNames\" lastCell=\"A2\")"})

	out := filepath.Join(outDir, "14_multisheet.xlsx")
	data := map[string]any{
		"sheetNames": []any{"Engineering", "Sales", "HR"},
		"departments": []any{
			map[string]any{"Name": "Engineering", "Head": "Alice"},
			map[string]any{"Name": "Sales", "Head": "Bob"},
			map[string]any{"Name": "HR", "Head": "Carol"},
		},
	}
	if err := fillTemplate(f, "t14.xlsx", out, data); err != nil {
		return err
	}
	of, err := excelize.OpenFile(out)
	if err != nil {
		return err
	}
	defer of.Close()
	sheets := of.GetSheetList()
	if len(sheets) < 3 {
		return fmt.Errorf("expected at least 3 sheets, got %d: %v", len(sheets), sheets)
	}
	return nil
}

// ===== 15: Custom Expression Notation =====
func testCustomNotation(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Name")
	f.SetCellValue(s, "A2", "{{e.Name}}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="items" var="e" lastCell="A2")`})

	out := filepath.Join(outDir, "15_custom_notation.xlsx")
	data := map[string]any{
		"items": []any{
			map[string]any{"Name": "Alpha"},
			map[string]any{"Name": "Beta"},
		},
	}
	if err := fillTemplate(f, "t15.xlsx", out, data, xlfill.WithExpressionNotation("{{", "}}")); err != nil {
		return err
	}
	if err := expectCell(out, s, "A2", "Alpha"); err != nil {
		return err
	}
	return expectCell(out, s, "A3", "Beta")
}

// ===== 16: Keep Template Sheet =====
func testKeepTemplateSheet(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Title")
	f.SetCellValue(s, "A2", "${e}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="items" var="e" lastCell="A2")`})

	out := filepath.Join(outDir, "16_keep_template.xlsx")
	data := map[string]any{"items": []any{"X", "Y"}}
	if err := fillTemplate(f, "t16.xlsx", out, data, xlfill.WithKeepTemplateSheet(true)); err != nil {
		return err
	}
	of, err := excelize.OpenFile(out)
	if err != nil {
		return err
	}
	defer of.Close()
	// Template sheet should still exist
	return nil
}

// ===== 17: AutoRowHeight =====
func testAutoRowHeight(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "${text}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: "jx:area(lastCell=\"A1\")\njx:autoRowHeight(lastCell=\"A1\")"})

	out := filepath.Join(outDir, "17_autorowheight.xlsx")
	data := map[string]any{"text": "This is a long text that should cause the row height to be adjusted automatically for better readability."}
	if err := fillTemplate(f, "t17.xlsx", out, data); err != nil {
		return err
	}
	return expectCell(out, s, "A1", data["text"].(string))
}

// ===== 18: FillBytes =====
func testFillBytes(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Value")
	f.SetCellValue(s, "A2", "${e}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="items" var="e" lastCell="A2")`})

	tmplPath := filepath.Join(inputDir, "t18.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		return err
	}
	f.Close()

	data := map[string]any{"items": []any{"One", "Two", "Three"}}
	outBytes, err := xlfill.FillBytes(tmplPath, data)
	if err != nil {
		return err
	}

	out := filepath.Join(outDir, "18_fill_bytes.xlsx")
	if err := os.WriteFile(out, outBytes, 0o644); err != nil {
		return err
	}

	return expectCell(out, s, "A3", "Two")
}

// ===== 19: FillReader =====
func testFillReader(outDir string) error {
	f := excelize.NewFile()
	s := "Sheet1"
	f.SetCellValue(s, "A1", "Item")
	f.SetCellValue(s, "A2", "${e}")

	f.AddComment(s, excelize.Comment{Cell: "A1", Author: "xlfill", Text: `jx:area(lastCell="A2")`})
	f.AddComment(s, excelize.Comment{Cell: "A2", Author: "xlfill", Text: `jx:each(items="items" var="e" lastCell="A2")`})

	// Save template to input/ for inspection, then read back as bytes
	tmplPath := filepath.Join(inputDir, "t19.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		return err
	}
	f.Close()

	tmplBytes, err := os.ReadFile(tmplPath)
	if err != nil {
		return err
	}

	data := map[string]any{"items": []any{"Red", "Green", "Blue"}}
	var outBuf bytes.Buffer
	if err := xlfill.FillReader(bytes.NewReader(tmplBytes), &outBuf, data); err != nil {
		return err
	}

	out := filepath.Join(outDir, "19_fill_reader.xlsx")
	if err := os.WriteFile(out, outBuf.Bytes(), 0o644); err != nil {
		return err
	}

	return expectCell(out, s, "A4", "Blue")
}
