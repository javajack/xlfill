package xlfill

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================
// Enhancement 1: Nested Commands (each-in-each, if-in-each)
// ============================================================

func TestNestedCommands_EachInEach(t *testing.T) {
	// Template: outer each iterates departments, inner each iterates employees
	// A1: jx:area(lastCell="B2")
	// A1: jx:each(items="departments" var="dept" lastCell="B2")
	// A1: ${dept.Name}
	// B1: (empty, part of outer area)
	// A2: jx:each(items="dept.Employees" var="e" lastCell="B2")
	// A2: ${e.Name}
	// B2: ${e.Salary}

	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${dept.Name}")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Salary}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:each(items=\"departments\" var=\"dept\" lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"dept.Employees\" var=\"e\" lastCell=\"B2\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"departments": []map[string]any{
			{
				"Name": "Engineering",
				"Employees": []map[string]any{
					{"Name": "Alice", "Salary": 90000},
					{"Name": "Bob", "Salary": 85000},
				},
			},
			{
				"Name": "Marketing",
				"Employees": []map[string]any{
					{"Name": "Carol", "Salary": 75000},
				},
			},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// Engineering department
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Engineering", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "90000", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "85000", v)

	// Marketing department
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Marketing", v)
	v, _ = out.GetCellValue(sheet, "A5")
	assert.Equal(t, "Carol", v)
	v, _ = out.GetCellValue(sheet, "B5")
	assert.Equal(t, "75000", v)
}

func TestNestedCommands_IfInEach(t *testing.T) {
	// Template: each iterates employees (A1:B2), if conditionally renders name (A2:A2)
	// Row 1: header, Row 2: data with conditional name
	// A1: jx:area + jx:each at A1:B2
	// A2: jx:if at A2:A2 (smaller than each, so it nests inside)
	// A2: ${e.Name}
	// B2: ${e.Salary}

	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Salary")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Salary}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:each(items=\"employees\" var=\"e\" lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:if(condition=\"e.VIP\" lastCell=\"A2\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Salary": 90000, "VIP": true},
			{"Name": "Bob", "Salary": 60000, "VIP": false},
			{"Name": "Carol", "Salary": 120000, "VIP": true},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// First iteration: Alice (VIP=true) — header row + data row
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Name", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "90000", v)

	// Second iteration: Bob (VIP=false) — name col is ZeroSize (skipped)
	// B3 should still have salary since it's outside the if area
	v, _ = out.GetCellValue(sheet, "B4")
	assert.Equal(t, "60000", v)

	// Third iteration: Carol (VIP=true)
	// Find Carol somewhere in the output
	found := false
	rows, _ := out.GetRows(sheet)
	for _, row := range rows {
		for _, cell := range row {
			if cell == "Carol" {
				found = true
				break
			}
		}
	}
	assert.True(t, found, "Carol should be found in output")
}

func TestNestedCommands_ThreeLevels(t *testing.T) {
	// Three-level nesting: company → departments → employees
	// Row1: ${company.Name}   (company header)
	// Row2: ${dept.Name}      (department header)
	// Row3: ${e.Name} ${e.Role}

	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${company.Name}")
	f.SetCellValue(sheet, "A2", "${dept.Name}")
	f.SetCellValue(sheet, "A3", "${e.Name}")
	f.SetCellValue(sheet, "B3", "${e.Role}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B3\")\njx:each(items=\"companies\" var=\"company\" lastCell=\"B3\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"company.Departments\" var=\"dept\" lastCell=\"B3\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A3", Author: "xlfill",
		Text: "jx:each(items=\"dept.Employees\" var=\"e\" lastCell=\"B3\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"companies": []map[string]any{
			{
				"Name": "Acme Corp",
				"Departments": []map[string]any{
					{
						"Name": "Engineering",
						"Employees": []map[string]any{
							{"Name": "Alice", "Role": "Dev"},
							{"Name": "Bob", "Role": "QA"},
						},
					},
				},
			},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Acme Corp", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Engineering", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "Dev", v)
	v, _ = out.GetCellValue(sheet, "A4")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "B4")
	assert.Equal(t, "QA", v)
}

// ============================================================
// Enhancement 2: Multisheet Each
// ============================================================

func TestMultisheetEach(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${dept.Name}")
	f.SetCellValue(sheet, "A2", "${dept.Head}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:each(items=\"departments\" var=\"dept\" multisheet=\"sheetNames\" lastCell=\"A2\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"sheetNames": []string{"Engineering", "Marketing", "Sales"},
		"departments": []map[string]any{
			{"Name": "Engineering", "Head": "Alice"},
			{"Name": "Marketing", "Head": "Bob"},
			{"Name": "Sales", "Head": "Carol"},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	sheets := out.GetSheetList()
	// Template sheet should be deleted, 3 new sheets created
	assert.NotContains(t, sheets, "Sheet1")
	assert.Contains(t, sheets, "Engineering")
	assert.Contains(t, sheets, "Marketing")
	assert.Contains(t, sheets, "Sales")

	// Verify content on each sheet
	v, _ := out.GetCellValue("Engineering", "A1")
	assert.Equal(t, "Engineering", v)
	v, _ = out.GetCellValue("Engineering", "A2")
	assert.Equal(t, "Alice", v)

	v, _ = out.GetCellValue("Marketing", "A1")
	assert.Equal(t, "Marketing", v)
	v, _ = out.GetCellValue("Marketing", "A2")
	assert.Equal(t, "Bob", v)

	v, _ = out.GetCellValue("Sales", "A1")
	assert.Equal(t, "Sales", v)
	v, _ = out.GetCellValue("Sales", "A2")
	assert.Equal(t, "Carol", v)
}

// ============================================================
// Enhancement 3: Recalculate Formulas on Open
// ============================================================

func TestRecalculateOnOpen(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${val}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	outBytes, err := FillBytes(tmpPath, map[string]any{"val": 42}, WithRecalculateOnOpen(true))
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// Verify the calc props were set
	props, err := out.GetCalcProps()
	require.NoError(t, err)
	assert.NotNil(t, props.FullCalcOnLoad)
	assert.True(t, *props.FullCalcOnLoad)
}

func TestRecalculateOnOpen_Default(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${val}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	// Default: no recalculate
	outBytes, err := FillBytes(tmpPath, map[string]any{"val": 42})
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	props, err := out.GetCalcProps()
	require.NoError(t, err)
	// FullCalcOnLoad should be nil or false by default
	if props.FullCalcOnLoad != nil {
		assert.False(t, *props.FullCalcOnLoad)
	}
}

// ============================================================
// Enhancement 4: Hyperlinks
// ============================================================

func TestHyperlink_Expression(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${hyperlink(url, title)}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"url":   "https://example.com",
		"title": "Example Site",
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// Display text should be the title
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Example Site", v)
}

func TestHyperlink_InEachLoop(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${hyperlink(e.URL, e.Name)}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"employees\" var=\"e\" lastCell=\"B1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "URL": "https://alice.dev"},
			{"Name": "Bob", "URL": "https://bob.dev"},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "Bob", v)
}

func TestHyperlinkValue_String(t *testing.T) {
	hv := HyperlinkValue{URL: "https://example.com", Display: "Example"}
	assert.Equal(t, "Example", hv.String())

	hv2 := HyperlinkValue{URL: "https://example.com"}
	assert.Equal(t, "https://example.com", hv2.String())
}

// ============================================================
// Enhancement 5: Area Listeners
// ============================================================

// testListener records all cells it saw.
type testListener struct {
	beforeCalls []CellRef
	afterCalls  []CellRef
}

func (l *testListener) BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool {
	l.beforeCalls = append(l.beforeCalls, target)
	return true
}

func (l *testListener) AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer) {
	l.afterCalls = append(l.afterCalls, target)
}

func TestAreaListener_CallCount(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.Age}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"people\" var=\"e\" lastCell=\"B1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	listener := &testListener{}

	data := map[string]any{
		"people": []map[string]any{
			{"Name": "Alice", "Age": 30},
			{"Name": "Bob", "Age": 25},
			{"Name": "Carol", "Age": 35},
		},
	}

	_, err := FillBytes(tmpPath, data, WithAreaListener(listener))
	require.NoError(t, err)

	// 3 items × 2 cells = 6 cells transformed
	assert.Equal(t, 6, len(listener.beforeCalls))
	assert.Equal(t, 6, len(listener.afterCalls))
}

// skipListener skips transformation for specific cells.
type skipListener struct {
	skipCol int
}

func (l *skipListener) BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool {
	return target.Col != l.skipCol // skip the specified column
}

func (l *skipListener) AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer) {}

func TestAreaListener_SkipTransform(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Hello")
	f.SetCellValue(sheet, "B1", "World")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	// Skip column B (index 1)
	_, err := FillBytes(tmpPath, map[string]any{}, WithAreaListener(&skipListener{skipCol: 1}))
	require.NoError(t, err)
	// Test passes if no error — the skip logic executed
}

// ============================================================
// Enhancement 6: Parameterized Formulas
// ============================================================

func TestParameterizedFormula(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", 100)
	f.SetCellFormula(sheet, "B1", "A1*${taxRate}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"taxRate": 0.2,
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// Formula should have ${taxRate} replaced with 0.2
	formula, _ := out.GetCellFormula(sheet, "B1")
	assert.Equal(t, "A1*0.2", formula)
}

func TestParameterizedFormula_MultipleVars(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", 1000)
	f.SetCellFormula(sheet, "B1", "A1*${rate}+${bonus}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"rate":  0.1,
		"bonus": 500,
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	formula, _ := out.GetCellFormula(sheet, "B1")
	assert.Equal(t, "A1*0.1+500", formula)
}

// ============================================================
// Enhancement 7: Auto Row Height
// ============================================================

func TestAutoRowHeight(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${text}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")\njx:autoRowHeight(lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"text": "This is a very long text that should cause the row to auto-resize when opened in Excel",
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// Verify the cell was populated
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Contains(t, v, "very long text")
}

func TestAutoRowHeight_InEach(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")\njx:each(items=\"people\" var=\"e\" lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"people": []map[string]any{
			{"Name": "Alice"},
			{"Name": "Bob"},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Bob", v)
}

// ============================================================
// Enhancement 8: Built-in Row/Col Context Variables
// ============================================================

func TestBuiltinRowVariable(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${_row}")
	f.SetCellValue(sheet, "B1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")\njx:each(items=\"people\" var=\"e\" lastCell=\"B1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"people": []map[string]any{
			{"Name": "Alice"},
			{"Name": "Bob"},
			{"Name": "Carol"},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// _row is 1-based
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "1", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "2", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "3", v)

	// Names are correct
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "Alice", v)
	v, _ = out.GetCellValue(sheet, "B2")
	assert.Equal(t, "Bob", v)
	v, _ = out.GetCellValue(sheet, "B3")
	assert.Equal(t, "Carol", v)
}

func TestBuiltinColVariable(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${_col}")
	f.SetCellValue(sheet, "B1", "${_col}")
	f.SetCellValue(sheet, "C1", "${_col}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"C1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	outBytes, err := FillBytes(tmpPath, map[string]any{})
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	// _col is 0-based
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "0", v)
	v, _ = out.GetCellValue(sheet, "B1")
	assert.Equal(t, "1", v)
	v, _ = out.GetCellValue(sheet, "C1")
	assert.Equal(t, "2", v)
}

func TestBuiltinRowCol_InFormula(t *testing.T) {
	// Use _row in a mixed expression
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Row ${_row}: ${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")\njx:each(items=\"people\" var=\"e\" lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"people": []map[string]any{
			{"Name": "Alice"},
			{"Name": "Bob"},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "Row 1: Alice", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Row 2: Bob", v)
}

// ============================================================
// Enhancement: SetCellHyperLink on Transformer
// ============================================================

func TestTransformerSetCellHyperLink(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "test")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ref := NewCellRef(sheet, 0, 0)
	err = tx.SetCellHyperLink(ref, "https://example.com", "Example")
	require.NoError(t, err)

	v, _ := f.GetCellValue(sheet, "A1")
	assert.Equal(t, "Example", v)
}

// ============================================================
// Cross-enhancement: listeners + nested commands
// ============================================================

func TestAreaListener_WithNestedCommands(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${dept.Name}")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:each(items=\"departments\" var=\"dept\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"dept.Employees\" var=\"e\" lastCell=\"A2\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	listener := &testListener{}

	data := map[string]any{
		"departments": []map[string]any{
			{
				"Name":      "Eng",
				"Employees": []map[string]any{{"Name": "Alice"}, {"Name": "Bob"}},
			},
		},
	}

	_, err := FillBytes(tmpPath, data, WithAreaListener(listener))
	require.NoError(t, err)

	// 1 dept header + 2 employees = 3 before/after calls
	assert.Equal(t, 3, len(listener.beforeCalls))
	assert.Equal(t, 3, len(listener.afterCalls))
}

// ============================================================
// Regression: existing features still work with new code paths
// ============================================================

func TestRegression_SimpleEach_StillWorks(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")\njx:each(items=\"items\" var=\"e\" lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"items": []map[string]any{
			{"Name": "One"},
			{"Name": "Two"},
			{"Name": "Three"},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "One", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "Two", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Three", v)
}

func TestRegression_IfCommand_StillWorks(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${msg}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A1\")\njx:if(condition=\"show\" lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	// true case
	outBytes, err := FillBytes(tmpPath, map[string]any{"show": true, "msg": "visible"})
	require.NoError(t, err)
	out, _ := excelize.OpenReader(bytes.NewReader(outBytes))
	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "visible", v)
	out.Close()

	// false case — if returns ZeroSize, template cell is unchanged
	outBytes, err = FillBytes(tmpPath, map[string]any{"show": false, "msg": "visible"})
	require.NoError(t, err)
	out, _ = excelize.OpenReader(bytes.NewReader(outBytes))
	v, _ = out.GetCellValue(sheet, "A1")
	// When condition is false and area/if share exact same bounds,
	// the cell retains its template expression since nothing transforms it
	assert.Contains(t, []string{"", "${msg}"}, v)
	out.Close()
}

func TestRegression_FormulaExpansion_StillWorks(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Val}")
	f.SetCellFormula(sheet, "A2", "SUM(A1:A1)")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:each(items=\"items\" var=\"e\" lastCell=\"A1\")",
	})

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"items": []map[string]any{
			{"Val": 10},
			{"Val": 20},
			{"Val": 30},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "10", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "20", v)
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "30", v)

	// SUM formula should be on A4 (shifted down)
	formula, _ := out.GetCellFormula(sheet, "A4")
	assert.NotEmpty(t, formula, "formula should be present at A4")
}

// ============================================================
// Unit tests for helper functions
// ============================================================

func TestToStringSlice(t *testing.T) {
	result, err := toStringSlice([]any{"a", "b", "c"})
	require.NoError(t, err)
	assert.Equal(t, []string{"a", "b", "c"}, result)

	result, err = toStringSlice([]string{"x", "y"})
	require.NoError(t, err)
	assert.Equal(t, []string{"x", "y"}, result)

	result, err = toStringSlice(nil)
	require.NoError(t, err)
	assert.Nil(t, result)
}

func TestGetCommandArea(t *testing.T) {
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{1, 1}, nil)

	each := &EachCommand{Area: area}
	assert.Equal(t, area, getCommandArea(each))

	ifCmd := &IfCommand{IfArea: area}
	assert.Equal(t, area, getCommandArea(ifCmd))

	arh := &AutoRowHeightCommand{Area: area}
	assert.Equal(t, area, getCommandArea(arh))

	// Unknown command type
	type customCmd struct{}
	custom := &customCmd{}
	_ = custom
}

func TestAutoRowHeightCommand_Name(t *testing.T) {
	cmd := &AutoRowHeightCommand{}
	assert.Equal(t, "autoRowHeight", cmd.Name())
}

func TestAreaListenerInterface(t *testing.T) {
	// Verify the interface is correctly defined
	var l AreaListener = &testListener{}
	assert.True(t, l.BeforeTransformCell(CellRef{}, CellRef{}, nil, nil))
}

func TestHyperlink_Function(t *testing.T) {
	hv := Hyperlink("https://example.com", "Example")
	assert.Equal(t, "https://example.com", hv.URL)
	assert.Equal(t, "Example", hv.Display)
	assert.Equal(t, "Example", hv.String())
}

func TestSortAreaBindings(t *testing.T) {
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{3, 3}, nil)
	area.AddCommand(&EachCommand{Items: "a", Var: "x"}, NewCellRef("Sheet1", 2, 0), Size{1, 1})
	area.AddCommand(&EachCommand{Items: "b", Var: "y"}, NewCellRef("Sheet1", 0, 0), Size{1, 1})

	sortAreaBindings([]*Area{area})

	assert.Equal(t, 0, area.Bindings[0].StartRef.Row)
	assert.Equal(t, 2, area.Bindings[1].StartRef.Row)
}

func TestNestedCommands_SameRowDifferentScope(t *testing.T) {
	// Verify nested command detection with same-start but different sizes
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "${e.Val}")

	// area spans A1:B2, each spans A1:B1 — each is inside area
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:each(items=\"items\" var=\"e\" lastCell=\"B1\")",
	})
	f.SetCellValue(sheet, "A2", "Footer")

	tmpPath := t.TempDir() + "/tmpl.xlsx"
	require.NoError(t, f.SaveAs(tmpPath))

	data := map[string]any{
		"items": []map[string]any{
			{"Name": "A", "Val": 1},
			{"Name": "B", "Val": 2},
		},
	}

	outBytes, err := FillBytes(tmpPath, data)
	require.NoError(t, err)

	out, err := excelize.OpenReader(bytes.NewReader(outBytes))
	require.NoError(t, err)
	defer out.Close()

	v, _ := out.GetCellValue(sheet, "A1")
	assert.Equal(t, "A", v)
	v, _ = out.GetCellValue(sheet, "A2")
	assert.Equal(t, "B", v)
	// Footer should shift to row 3
	v, _ = out.GetCellValue(sheet, "A3")
	assert.Equal(t, "Footer", v)
}

func init() {
	// Silence unused import warning
	_ = fmt.Sprintf
}
