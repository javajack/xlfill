package goxls

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// testdataDir returns the path to testdata directory, creating it if needed.
func testdataDir(t *testing.T) string {
	t.Helper()
	dir := filepath.Join("testdata")
	require.NoError(t, os.MkdirAll(dir, 0o755))
	return dir
}

// createBasicTemplate creates a basic test template with various cell types and comments.
// Layout:
//
//	A1: "Name" (bold)     B1: "Age"     C1: "Salary"
//	A2: "${e.Name}"       B2: "${e.Age}" C2: "${e.Salary}"
//
// A1 has comment: jx:area(lastCell="C2")
// A2 has comment: jx:each(items="employees" var="e" lastCell="C2")
func createBasicTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create bold style for header
	boldStyle, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
	})
	require.NoError(t, err)

	// Header row
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Age")
	f.SetCellValue(sheet, "C1", "Salary")
	f.SetCellStyle(sheet, "A1", "C1", boldStyle)

	// Data row with expressions
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Age}")
	f.SetCellValue(sheet, "C2", "${e.Salary}")

	// Comments (JXLS commands)
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

	path := filepath.Join(testdataDir(t), "basic_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createStyledTemplate creates a template with various styles to test preservation.
func createStyledTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Style with red fill + bold
	redBold, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Color: "FF0000"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"FFEEEE"}, Pattern: 1},
	})
	require.NoError(t, err)

	// Number format style
	numFmt, err := f.NewStyle(&excelize.Style{
		NumFmt: 4, // #,##0.00
	})
	require.NoError(t, err)

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellStyle(sheet, "A1", "A1", redBold)

	f.SetCellValue(sheet, "B1", 1234.5)
	f.SetCellStyle(sheet, "B1", "B1", numFmt)

	f.SetCellValue(sheet, "C1", "Plain")

	path := filepath.Join(testdataDir(t), "styled_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createFormulaTemplate creates a template with formulas.
func createFormulaTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Value")
	f.SetCellValue(sheet, "A2", "${e.Amount}")
	f.SetCellFormula(sheet, "A3", "SUM(A2:A2)")

	f.AddComment(sheet, excelize.Comment{
		Cell:   "A1",
		Author: "goxls",
		Text:   `jx:area(lastCell="A3")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell:   "A2",
		Author: "goxls",
		Text:   `jx:each(items="items" var="e" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "formula_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createMergedTemplate creates a template with merged cells.
func createMergedTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Merged Header")
	f.MergeCell(sheet, "A1", "C1")

	f.SetCellValue(sheet, "A2", "Col1")
	f.SetCellValue(sheet, "B2", "Col2")
	f.SetCellValue(sheet, "C2", "Col3")

	path := filepath.Join(testdataDir(t), "merged_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}
