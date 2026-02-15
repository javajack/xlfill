package xlfill

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

func TestDescribe_BasicTemplate(t *testing.T) {
	tmpl := createValidTemplate(t)
	output, err := Describe(tmpl)
	require.NoError(t, err)

	assert.Contains(t, output, "Template:")
	assert.Contains(t, output, "area (2x2)")
	assert.Contains(t, output, "each")
	assert.Contains(t, output, `items="employees"`)
	assert.Contains(t, output, `var="e"`)
	assert.Contains(t, output, "${e.Name}")
	assert.Contains(t, output, "${e.Age}")
}

func TestDescribe_NestedCommands(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := "Sheet1"

	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Status")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Status}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="B2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "B2", Author: "xlfill",
		Text: `jx:if(condition="e.Active" lastCell="B2")`,
	})

	path := filepath.Join(testdataDir(t), "describe_nested.xlsx")
	require.NoError(t, f.SaveAs(path))
	t.Cleanup(func() { os.Remove(path) })

	output, err := Describe(path)
	require.NoError(t, err)

	assert.Contains(t, output, "each")
	assert.Contains(t, output, "if")
	assert.Contains(t, output, `condition="e.Active"`)
	// The if command should be indented more than the each command
	eachIdx := indexOf(output, "each")
	ifIdx := indexOf(output, "if")
	assert.Greater(t, ifIdx, eachIdx, "if command should appear after each command in nested tree")
}

func TestDescribe_BadTemplatePath(t *testing.T) {
	output, err := Describe("/nonexistent/template.xlsx")
	assert.Error(t, err)
	assert.Empty(t, output)
}

func TestDescribe_TopLevelFunction(t *testing.T) {
	tmpl := createValidTemplate(t)

	// Top-level function
	output1, err1 := Describe(tmpl)
	require.NoError(t, err1)

	// Via NewFiller
	filler := NewFiller(WithTemplate(tmpl))
	output2, err2 := filler.Describe()
	require.NoError(t, err2)

	assert.Equal(t, output1, output2)
}

// indexOf returns the position of the first occurrence of substr in s, or -1.
func indexOf(s, substr string) int {
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return i
		}
	}
	return -1
}
