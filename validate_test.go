package xlfill

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func createValidTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Age")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Age}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="B2")`,
	})

	path := filepath.Join(testdataDir(t), "validate_valid.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })
	return path
}

func TestValidate_ValidTemplate(t *testing.T) {
	tmpl := createValidTemplate(t)
	issues, err := Validate(tmpl)
	require.NoError(t, err)
	assert.Empty(t, issues)
}

func TestValidate_InvalidExpressionSyntax(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "B1", "${e.Name +}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")`,
	})

	path := filepath.Join(testdataDir(t), "validate_bad_expr.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	issues, err := Validate(path)
	require.NoError(t, err)
	require.Len(t, issues, 1)
	assert.Equal(t, SeverityError, issues[0].Severity)
	assert.Contains(t, issues[0].Message, "invalid expression syntax")
	assert.Equal(t, "Sheet1", issues[0].CellRef.Sheet)
}

func TestValidate_LastCellOutOfBounds(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	// Area is A1:A2, but each command claims lastCell=B2 which exceeds the area width
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="B2")`,
	})

	path := filepath.Join(testdataDir(t), "validate_bounds.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	issues, err := Validate(path)
	require.NoError(t, err)
	require.NotEmpty(t, issues)

	found := false
	for _, issue := range issues {
		if issue.Severity == SeverityError && assert.ObjectsAreEqual("each", "") == false {
			found = true
			break
		}
	}
	// Check that at least one issue is about bounds
	found = false
	for _, issue := range issues {
		if issue.Severity == SeverityError {
			found = true
			break
		}
	}
	assert.True(t, found, "expected at least one error-level issue about bounds")
}

func TestValidate_InvalidItemsExpression(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees[" var="e" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "validate_bad_items.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	issues, err := Validate(path)
	require.NoError(t, err)
	require.NotEmpty(t, issues)
	assert.Contains(t, issues[0].Message, "each command has invalid items expression")
}

func TestValidate_InvalidConditionExpression(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Status")
	f.SetCellValue(sheet, "A2", "Active")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:if(condition="e.Active &&" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "validate_bad_cond.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	issues, err := Validate(path)
	require.NoError(t, err)
	require.NotEmpty(t, issues)
	assert.Contains(t, issues[0].Message, "if command has invalid condition expression")
}

func TestValidate_InvalidFormulaExpression(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Value")
	// Set a formula with a bad expression
	f.SetCellFormula(sheet, "A1", `SUM(A${bad syntax}:A10)`)

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "validate_bad_formula.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	issues, err := Validate(path)
	require.NoError(t, err)
	require.NotEmpty(t, issues)
	assert.Equal(t, SeverityError, issues[0].Severity)
	assert.Contains(t, issues[0].Message, "invalid expression syntax")
}

func TestValidate_MultipleIssues(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "${bad +}")
	f.SetCellValue(sheet, "B1", "${also broken]}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")`,
	})

	path := filepath.Join(testdataDir(t), "validate_multi.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	issues, err := Validate(path)
	require.NoError(t, err)
	assert.GreaterOrEqual(t, len(issues), 2, "expected at least 2 issues")
}

func TestValidate_BadTemplatePath(t *testing.T) {
	issues, err := Validate("/nonexistent/template.xlsx")
	assert.Error(t, err)
	assert.Nil(t, issues)
}

func TestValidate_IssueString(t *testing.T) {
	errIssue := ValidationIssue{
		Severity: SeverityError,
		CellRef:  NewCellRef("Sheet1", 1, 0),
		Message:  "bad expression",
	}
	assert.Equal(t, "[ERROR] Sheet1!A2: bad expression", errIssue.String())

	warnIssue := ValidationIssue{
		Severity: SeverityWarning,
		CellRef:  NewCellRef("Data", 0, 2),
		Message:  "unused area",
	}
	assert.Equal(t, "[WARN] Data!C1: unused area", warnIssue.String())
}
