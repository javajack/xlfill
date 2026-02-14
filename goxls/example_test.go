package goxls_test

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"

	"github.com/mhseiden/goxls"
	"github.com/xuri/excelize/v2"
)

func ExampleFill() {
	// Create a template programmatically (normally you'd use an existing .xlsx file)
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Header row
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Age")
	f.SetCellValue(sheet, "C1", "Salary")

	// Data row with expressions
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Age}")
	f.SetCellValue(sheet, "C2", "${e.Salary}")

	// JXLS-style commands in cell comments
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A1",
		Author: "goxls",
		Text:   `jx:area(lastCell="C2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A2",
		Author: "goxls",
		Text:   `jx:each(items="employees" var="e" lastCell="C2")`,
	})

	// Save template
	tmpDir := os.TempDir()
	tmplPath := filepath.Join(tmpDir, "example_template.xlsx")
	f.SaveAs(tmplPath)
	f.Close()
	defer os.Remove(tmplPath)

	// Fill with data
	data := map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Age": 30, "Salary": 5000},
			map[string]any{"Name": "Bob", "Age": 25, "Salary": 6000},
			map[string]any{"Name": "Carol", "Age": 35, "Salary": 7000},
		},
	}

	outPath := filepath.Join(tmpDir, "example_output.xlsx")
	defer os.Remove(outPath)

	err := goxls.Fill(tmplPath, outPath, data)
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	// Read output to verify
	out, _ := excelize.OpenFile(outPath)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A2")
	fmt.Println(v)
	v, _ = out.GetCellValue(sheet, "A3")
	fmt.Println(v)
	v, _ = out.GetCellValue(sheet, "A4")
	fmt.Println(v)
	// Output:
	// Alice
	// Bob
	// Carol
}

func ExampleFillBytes() {
	// Create template
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Item")
	f.SetCellValue(sheet, "A2", "${e}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "goxls",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "goxls",
		Text: `jx:each(items="items" var="e" lastCell="A2")`,
	})

	tmpPath := filepath.Join(os.TempDir(), "example_bytes.xlsx")
	f.SaveAs(tmpPath)
	f.Close()
	defer os.Remove(tmpPath)

	// Get output as bytes
	outBytes, err := goxls.FillBytes(tmpPath, map[string]any{
		"items": []any{"Apple", "Banana", "Cherry"},
	})
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	// Read from bytes
	out, _ := excelize.OpenReader(bytes.NewReader(outBytes))
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A2")
	fmt.Println(v)
	// Output:
	// Apple
}
