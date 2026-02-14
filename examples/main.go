package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/javajack/xlfill"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Step 1: Create a template programmatically.
	// In real use, you'd design this in Excel/LibreOffice with formatting, colors, etc.
	tmplPath := filepath.Join(os.TempDir(), "template.xlsx")
	createTemplate(tmplPath)

	// Step 2: Prepare data.
	data := map[string]any{
		"title": "Employee Report",
		"employees": []map[string]any{
			{"Name": "Alice", "Department": "Engineering", "Salary": 95000},
			{"Name": "Bob", "Department": "Marketing", "Salary": 72000},
			{"Name": "Carol", "Department": "Engineering", "Salary": 105000},
			{"Name": "David", "Department": "Sales", "Salary": 68000},
			{"Name": "Eve", "Department": "Marketing", "Salary": 81000},
		},
	}

	// Step 3: Fill the template.
	outPath := "output.xlsx"
	err := xlfill.Fill(tmplPath, outPath, data)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("Wrote %s\n", outPath)

	// Step 4: Verify by reading back.
	verifyOutput(outPath)
}

func createTemplate(path string) {
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Row 1: Title
	f.SetCellValue(sheet, "A1", "${title}")

	// Row 2: Headers
	f.SetCellValue(sheet, "A2", "Name")
	f.SetCellValue(sheet, "B2", "Department")
	f.SetCellValue(sheet, "C2", "Salary")

	// Row 3: Template row with expressions (will be repeated for each employee)
	f.SetCellValue(sheet, "A3", "${e.Name}")
	f.SetCellValue(sheet, "B3", "${e.Department}")
	f.SetCellValue(sheet, "C3", "${e.Salary}")

	// Add jx: commands as cell comments.
	// The area command marks A1:C3 as the working region.
	// The each command on A3 repeats row 3 for each employee.
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A1",
		Author: "xlfill",
		Text:   `jx:area(lastCell="C3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A3",
		Author: "xlfill",
		Text:   `jx:each(items="employees" var="e" lastCell="C3")`,
	})

	if err := f.SaveAs(path); err != nil {
		fmt.Fprintf(os.Stderr, "Error creating template: %v\n", err)
		os.Exit(1)
	}
}

func verifyOutput(path string) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error reading output: %v\n", err)
		return
	}
	defer f.Close()

	sheet := "Sheet1"
	rows, _ := f.GetRows(sheet)
	fmt.Println()
	for i, row := range rows {
		fmt.Printf("Row %d: %v\n", i+1, row)
	}
}
