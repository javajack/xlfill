package xlfill

import (
	"bytes"
	"context"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================================
// Enhancement 1: Differential Context Map Updates (Performance)
// ============================================================================

func TestContextDifferentialMapUpdate(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 1, "y": 2})

	// First call builds full map
	m1 := ctx.ToMap()
	assert.Equal(t, 1, m1["x"])
	assert.Equal(t, 2, m1["y"])

	// Set a runVar — should dirty the key, not rebuild
	ctx.setRunVar("e", "hello")
	m2 := ctx.ToMap()
	assert.Equal(t, "hello", m2["e"])
	assert.Equal(t, 1, m2["x"]) // data still present

	// Verify differential update: m1 and m2 should be same map
	// We verify by adding a sentinel to m1 and checking m2 sees it
	m1["_sentinel"] = true
	assert.Equal(t, true, m2["_sentinel"]) // same map
	delete(m1, "_sentinel")

	// Change the same runVar
	ctx.setRunVar("e", "world")
	m3 := ctx.ToMap()
	assert.Equal(t, "world", m3["e"])

	// Remove a runVar — should restore from data or delete
	ctx.removeRunVar("e")
	m4 := ctx.ToMap()
	_, hasE := m4["e"]
	assert.False(t, hasE) // "e" not in data, so removed from map

	// PutVar triggers full rebuild
	ctx.PutVar("z", 3)
	m5 := ctx.ToMap()
	assert.Equal(t, 3, m5["z"])
}

func TestContextDifferentialMapWithNestedLoops(t *testing.T) {
	ctx := NewContext(map[string]any{"data": "base"})

	// Simulate nested loop: outer sets "dept", inner sets "emp"
	ctx.setRunVar("dept", "Engineering")
	ctx.setRunVar("emp", "Alice")

	m := ctx.ToMap()
	assert.Equal(t, "Engineering", m["dept"])
	assert.Equal(t, "Alice", m["emp"])

	// Inner loop iteration changes only "emp"
	ctx.setRunVar("emp", "Bob")
	m = ctx.ToMap()
	assert.Equal(t, "Engineering", m["dept"]) // unchanged
	assert.Equal(t, "Bob", m["emp"])           // updated

	// Outer loop iteration changes "dept", removes "emp"
	ctx.removeRunVar("emp")
	ctx.setRunVar("dept", "Marketing")
	m = ctx.ToMap()
	assert.Equal(t, "Marketing", m["dept"])
	_, hasEmp := m["emp"]
	assert.False(t, hasEmp)
}

// ============================================================================
// Enhancement 3: Pre-allocated Slices (Structural)
// ============================================================================

func TestPreAllocatedCommentedCells(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	tx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer tx.Close()

	// Should return pre-allocated slice with correct items
	commented := tx.GetCommentedCells()
	assert.True(t, len(commented) > 0)

	formulas := tx.GetFormulaCells()
	// basic template has no formulas
	assert.Equal(t, 0, len(formulas))
}

// ============================================================================
// Enhancement 4: Structured Error Types
// ============================================================================

func TestStructuredErrorTypes(t *testing.T) {
	// Template error
	err := NewTemplateError(
		NewCellRef("Sheet1", 0, 0),
		"each",
		"missing items attribute",
		nil,
	)
	assert.Contains(t, err.Error(), "[xlfill:template]")
	assert.Contains(t, err.Error(), "Sheet1!A1")
	assert.Contains(t, err.Error(), "each")

	var xlErr *XLFillError
	assert.True(t, errors.As(err, &xlErr))
	assert.Equal(t, ErrTemplate, xlErr.Kind)

	// Data error
	err2 := NewDataError(
		NewCellRef("Sheet1", 1, 1),
		"variable 'e' not found",
		nil,
	)
	assert.Contains(t, err2.Error(), "[xlfill:data]")

	// Runtime error with wrapped cause
	inner := errors.New("file not found")
	err3 := NewRuntimeError("open template", inner)
	assert.Contains(t, err3.Error(), "[xlfill:runtime]")
	assert.True(t, errors.Is(err3, inner))
}

func TestErrorKindString(t *testing.T) {
	assert.Equal(t, "template", ErrTemplate.String())
	assert.Equal(t, "data", ErrData.String())
	assert.Equal(t, "runtime", ErrRuntime.String())
}

// ============================================================================
// Enhancement 5: Unknown Command Warnings (Strict Mode)
// ============================================================================

func createTemplateWithUnknownCommand(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Data")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:eache(items="employees" var="e" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "unknown_cmd_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func TestUnknownCommandWarning(t *testing.T) {
	tmpl := createTemplateWithUnknownCommand(t)
	defer os.Remove(tmpl)

	filler := NewFiller(WithTemplate(tmpl))
	tx, err := filler.openTemplate()
	require.NoError(t, err)
	defer tx.Close()

	_, err = filler.BuildAreas(tx)
	require.NoError(t, err) // default mode: no error

	warnings := filler.Warnings()
	require.Len(t, warnings, 1)
	assert.Contains(t, warnings[0].Message, `unknown command "eache"`)
	assert.Contains(t, warnings[0].Message, `did you mean "each"`)
}

func TestUnknownCommandStrictMode(t *testing.T) {
	tmpl := createTemplateWithUnknownCommand(t)
	defer os.Remove(tmpl)

	filler := NewFiller(WithTemplate(tmpl), WithStrictMode(true))
	tx, err := filler.openTemplate()
	require.NoError(t, err)
	defer tx.Close()

	_, err = filler.BuildAreas(tx)
	require.Error(t, err)
	assert.Contains(t, err.Error(), `unknown command "eache"`)
	assert.Contains(t, err.Error(), `did you mean "each"`)
}

func TestSuggestCommand(t *testing.T) {
	known := []string{"each", "if", "grid", "image", "mergeCells", "updateCell", "autoRowHeight", "repeat"}

	assert.Contains(t, suggestCommand("eache", known), `"each"`)
	assert.Contains(t, suggestCommand("iif", known), `"if"`)
	assert.Contains(t, suggestCommand("reapeat", known), `"repeat"`)
	assert.Equal(t, "", suggestCommand("totallyunknown", known)) // too different
}

func TestLevenshtein(t *testing.T) {
	assert.Equal(t, 0, levenshtein("abc", "abc"))
	assert.Equal(t, 1, levenshtein("abc", "ab"))
	assert.Equal(t, 1, levenshtein("abc", "abcd"))
	assert.Equal(t, 3, levenshtein("abc", "xyz"))
}

// ============================================================================
// Enhancement 6: Context Variable Scope Validation (via ValidateData)
// ============================================================================

func TestValidateDataMissingVariable(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	// data has "employees" — which is what jx:each(items="employees") needs
	// But expressions use ${e.Name}, and "e" is provided by the each command
	data := map[string]any{
		"employees": []map[string]any{{"Name": "Alice"}},
	}
	issues, err := ValidateData(tmpl, data)
	require.NoError(t, err)
	// "e" is a command var so should not be flagged
	for _, issue := range issues {
		assert.NotContains(t, issue.Message, `variable "e"`)
	}
}

func TestValidateDataMissingTopLevelVar(t *testing.T) {
	// Create template referencing a variable not in data
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${companyName}")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "validate_data_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	// data is missing "companyName"
	data := map[string]any{
		"employees": []map[string]any{{"Name": "Alice"}},
	}
	issues, err := ValidateData(path, data)
	require.NoError(t, err)

	foundMissing := false
	for _, issue := range issues {
		if strings.Contains(issue.Message, "companyName") {
			foundMissing = true
			break
		}
	}
	assert.True(t, foundMissing, "expected warning about missing 'companyName', got: %v", issues)
}

// ============================================================================
// Enhancement 7: Progress Reporting + Context Cancellation
// ============================================================================

func TestProgressReporting(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	var progressCalls []FillProgress
	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
		},
	}

	outPath := filepath.Join(testdataDir(t), "progress_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, data,
		WithProgressFunc(func(p FillProgress) {
			progressCalls = append(progressCalls, p)
		}),
	)
	require.NoError(t, err)
	assert.True(t, len(progressCalls) > 0, "expected progress callbacks")
}

func TestContextCancellation(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	ctx, cancel := context.WithCancel(context.Background())
	cancel() // cancel immediately

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	outPath := filepath.Join(testdataDir(t), "cancelled_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, data,
		WithContext(ctx),
	)
	assert.Error(t, err)
	assert.True(t, errors.Is(err, context.Canceled))
}

func TestContextTimeout(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	// Use an already-expired context for deterministic behavior
	ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(-time.Second))
	defer cancel()

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	outPath := filepath.Join(testdataDir(t), "timeout_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, data, WithContext(ctx))
	assert.Error(t, err)
}

// ============================================================================
// Enhancement 8: Data Contract Validation
// ============================================================================

func TestExtractTopLevelVars(t *testing.T) {
	vars := extractTopLevelVars("e.Name + company")
	assert.Contains(t, vars, "e")
	assert.Contains(t, vars, "company")

	vars2 := extractTopLevelVars("items[0].Name")
	assert.Contains(t, vars2, "items")

	vars3 := extractTopLevelVars("true && false")
	assert.Empty(t, vars3) // keywords filtered

	vars4 := extractTopLevelVars("len(items) > 0")
	assert.Contains(t, vars4, "items")
	assert.NotContains(t, vars4, "len") // keyword
}

func TestIsExprKeyword(t *testing.T) {
	assert.True(t, isExprKeyword("true"))
	assert.True(t, isExprKeyword("len"))
	assert.True(t, isExprKeyword("filter"))
	assert.False(t, isExprKeyword("employees"))
	assert.False(t, isExprKeyword("companyName"))
}

// ============================================================================
// Enhancement 9: Debug/Trace Mode
// ============================================================================

func TestDebugTrace(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	var buf bytes.Buffer
	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
		},
	}

	outPath := filepath.Join(testdataDir(t), "debug_out.xlsx")
	err := Fill(tmpl, outPath, data, WithDebugWriter(&buf))
	require.NoError(t, err)
	defer os.Remove(outPath)

	trace := buf.String()
	assert.Contains(t, trace, "[area]")
	assert.Contains(t, trace, "[each]")
	assert.Contains(t, trace, "[iter 0]")
	assert.Contains(t, trace, "[iter 1]")
	assert.Contains(t, trace, "[done]")
	assert.Contains(t, trace, "cells transformed")
}

func TestDebugTracerImplementsAreaListener(t *testing.T) {
	var buf bytes.Buffer
	tracer := NewDebugTracer(&buf)

	// Verify it implements AreaListener
	var _ AreaListener = tracer

	src := NewCellRef("Sheet1", 0, 0)
	target := NewCellRef("Sheet1", 0, 0)
	ctx := NewContext(nil)

	// BeforeTransformCell should return true (proceed)
	assert.True(t, tracer.BeforeTransformCell(src, target, ctx, nil))
	assert.Equal(t, int64(1), tracer.cellCount.Load())
}

// ============================================================================
// Enhancement 10: Compiled Template Reuse
// ============================================================================

func TestCompiledTemplate(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl)
	require.NoError(t, err)

	// Fill with first data set
	data1 := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}
	bytes1, err := compiled.FillBytes(data1)
	require.NoError(t, err)
	assert.True(t, len(bytes1) > 0)

	// Fill with second data set — reusing the compiled template
	data2 := map[string]any{
		"employees": []map[string]any{
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
		},
	}
	bytes2, err := compiled.FillBytes(data2)
	require.NoError(t, err)
	assert.True(t, len(bytes2) > 0)

	// Verify they produced different outputs (different data)
	assert.NotEqual(t, bytes1, bytes2)

	// Verify second output has correct data
	f, err := excelize.OpenReader(bytes.NewReader(bytes2))
	require.NoError(t, err)
	defer f.Close()

	val, err := f.GetCellValue("Sheet1", "A2")
	require.NoError(t, err)
	assert.Equal(t, "Bob", val)

	val2, err := f.GetCellValue("Sheet1", "A3")
	require.NoError(t, err)
	assert.Equal(t, "Carol", val2)
}

func TestCompiledTemplateToFile(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl)
	require.NoError(t, err)

	outPath := filepath.Join(testdataDir(t), "compiled_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}
	err = compiled.Fill(data, outPath)
	require.NoError(t, err)

	// Verify file exists and is valid
	_, err = os.Stat(outPath)
	require.NoError(t, err)
}

func TestCompileInvalidTemplate(t *testing.T) {
	_, err := Compile("nonexistent.xlsx")
	assert.Error(t, err)
}

// ============================================================================
// Enhancement 11: Style Callbacks (StyleListener)
// ============================================================================

type testStyleListener struct {
	calls int
}

func (l *testStyleListener) BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool {
	return true
}

func (l *testStyleListener) AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer) {}

func (l *testStyleListener) StyleCell(target CellRef, value any, ctx *Context) *StyleOverride {
	l.calls++
	// Make numbers > 1500 bold with red font
	if f, ok := toFloat64(value); ok && f > 1500 {
		bold := true
		color := "FF0000"
		return &StyleOverride{Bold: &bold, FontColor: &color}
	}
	return nil
}

func TestStyleListener(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	listener := &testStyleListener{}
	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
		},
	}

	outPath := filepath.Join(testdataDir(t), "style_listener_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, data, WithAreaListener(listener))
	require.NoError(t, err)
	assert.True(t, listener.calls > 0, "style listener should have been called")
}

// ============================================================================
// Enhancement 14: jx:repeat Command
// ============================================================================

func createRepeatTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Row ${i}")
	f.SetCellValue(sheet, "B1", "Data")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")
jx:repeat(count="3" var="i" lastCell="B1")`,
	})

	path := filepath.Join(testdataDir(t), "repeat_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func TestRepeatCommand(t *testing.T) {
	tmpl := createRepeatTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "repeat_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, map[string]any{})
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	// Should have 3 rows
	v1, _ := f.GetCellValue("Sheet1", "A1")
	v2, _ := f.GetCellValue("Sheet1", "A2")
	v3, _ := f.GetCellValue("Sheet1", "A3")

	assert.Equal(t, "Row 0", v1)
	assert.Equal(t, "Row 1", v2)
	assert.Equal(t, "Row 2", v3)
}

func TestRepeatCommandWithExpression(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Line ${i}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:repeat(count="numRows" var="i" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "repeat_expr_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "repeat_expr_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(path, outPath, map[string]any{"numRows": 5})
	require.NoError(t, err)

	out, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer out.Close()

	v5, _ := out.GetCellValue("Sheet1", "A5")
	assert.Equal(t, "Line 4", v5)
}

func TestRepeatDirectionRight(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Col ${i}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:repeat(count="4" var="i" direction="RIGHT" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "repeat_right_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "repeat_right_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(path, outPath, map[string]any{})
	require.NoError(t, err)

	out, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer out.Close()

	v1, _ := out.GetCellValue("Sheet1", "A1")
	v2, _ := out.GetCellValue("Sheet1", "B1")
	v3, _ := out.GetCellValue("Sheet1", "C1")
	v4, _ := out.GetCellValue("Sheet1", "D1")

	assert.Equal(t, "Col 0", v1)
	assert.Equal(t, "Col 1", v2)
	assert.Equal(t, "Col 2", v3)
	assert.Equal(t, "Col 3", v4)
}

func TestRepeatZeroCount(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Data")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:repeat(count="0" var="i" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "repeat_zero_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "repeat_zero_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(path, outPath, map[string]any{})
	require.NoError(t, err) // zero repeats should succeed silently
}

// ============================================================================
// Enhancement 15: Transformer Interface Split (CellReader/CellWriter/SheetManager)
// ============================================================================

func TestTransformerInterfaceComposition(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	tx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer tx.Close()

	// Verify ExcelizeTransformer satisfies all sub-interfaces
	var _ CellReader = tx
	var _ CellWriter = tx
	var _ SheetManager = tx
	var _ Transformer = tx
}

// ============================================================================
// Warning System
// ============================================================================

func TestWarningCollector(t *testing.T) {
	wc := &WarningCollector{}

	wc.Add(NewCellRef("Sheet1", 0, 0), "test warning 1")
	wc.Add(NewCellRef("Sheet1", 1, 0), "test warning 2")

	warnings := wc.Warnings()
	assert.Len(t, warnings, 2)
	assert.Equal(t, "test warning 1", warnings[0].Message)

	wc.Reset()
	assert.Len(t, wc.Warnings(), 0)
}

func TestWarningString(t *testing.T) {
	w := Warning{Cell: NewCellRef("Sheet1", 0, 0), Message: "hello"}
	assert.Equal(t, "[WARN] Sheet1!A1: hello", w.String())

	w2 := Warning{Message: "no cell"}
	assert.Equal(t, "[WARN] no cell", w2.String())
}

// ============================================================================
// Command Registry
// ============================================================================

func TestCommandRegistryKnownNames(t *testing.T) {
	reg := NewCommandRegistry()
	names := reg.KnownNames()

	assert.Contains(t, names, "each")
	assert.Contains(t, names, "if")
	assert.Contains(t, names, "grid")
	assert.Contains(t, names, "image")
	assert.Contains(t, names, "mergeCells")
	assert.Contains(t, names, "updateCell")
	assert.Contains(t, names, "autoRowHeight")
	assert.Contains(t, names, "repeat")
}

// ============================================================================
// Repeat Command Unit Tests
// ============================================================================

func TestRepeatCommandFromAttrs(t *testing.T) {
	cmd, err := newRepeatCommandFromAttrs(map[string]string{
		"count": "5",
		"var":   "i",
	})
	require.NoError(t, err)
	rc := cmd.(*RepeatCommand)
	assert.Equal(t, "5", rc.Count)
	assert.Equal(t, "i", rc.Var)
	assert.Equal(t, "DOWN", rc.Direction)
	assert.Equal(t, "repeat", rc.Name())
}

func TestRepeatCommandMissingCount(t *testing.T) {
	_, err := newRepeatCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "requires 'count'")
}

func TestRepeatCommandNoArea(t *testing.T) {
	cmd := &RepeatCommand{Count: "3", Var: "i", Direction: "DOWN"}
	ctx := NewContext(nil)
	_, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, nil)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no area")
}

// ============================================================================
// Describe with Repeat Command
// ============================================================================

func TestDescribeWithRepeat(t *testing.T) {
	tmpl := createRepeatTemplate(t)
	defer os.Remove(tmpl)

	desc, err := Describe(tmpl)
	require.NoError(t, err)
	assert.Contains(t, desc, "repeat")
	assert.Contains(t, desc, `count="3"`)
}

// ============================================================================
// Validate with Repeat Command
// ============================================================================

func TestValidateWithRepeatCommand(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${i}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:repeat(count="numRows" var="i" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "validate_repeat_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	issues, err := Validate(path)
	require.NoError(t, err)
	// Should have no errors — all expressions are valid
	for _, issue := range issues {
		assert.NotEqual(t, SeverityError, issue.Severity, "unexpected error: %s", issue.Message)
	}
}

// ============================================================================
// Integration: Enhanced Error Messages
// ============================================================================

func TestEnhancedErrorOnMissingTemplate(t *testing.T) {
	err := Fill("nonexistent.xlsx", "out.xlsx", nil)
	assert.Error(t, err)

	var xlErr *XLFillError
	if errors.As(err, &xlErr) {
		assert.Equal(t, ErrRuntime, xlErr.Kind)
	}
}

// ============================================================================
// Bug Fix: toFloat64 missing unsigned int types
// ============================================================================

func TestToFloat64AllTypes(t *testing.T) {
	tests := []struct {
		input    any
		expected float64
		ok       bool
	}{
		{int(5), 5.0, true},
		{int8(5), 5.0, true},
		{int16(5), 5.0, true},
		{int32(5), 5.0, true},
		{int64(5), 5.0, true},
		{uint(5), 5.0, true},
		{uint8(5), 5.0, true},
		{uint16(5), 5.0, true},
		{uint32(5), 5.0, true},
		{uint64(5), 5.0, true},
		{float32(5.5), 5.5, true},
		{float64(5.5), 5.5, true},
		{"not a number", 0, false},
		{nil, 0, false},
	}

	for _, tt := range tests {
		f, ok := toFloat64(tt.input)
		assert.Equal(t, tt.ok, ok, "toFloat64(%T)", tt.input)
		if ok {
			assert.InDelta(t, tt.expected, f, 0.01, "toFloat64(%T)", tt.input)
		}
	}
}

// ============================================================================
// Bug Fix: parseOrderBy with empty tokens
// ============================================================================

func TestParseOrderByEdgeCases(t *testing.T) {
	// Empty string
	specs := parseOrderBy("", "e")
	assert.Nil(t, specs)

	// Whitespace only
	specs = parseOrderBy("   ", "e")
	assert.Nil(t, specs)

	// Trailing comma
	specs = parseOrderBy("e.Name ASC,", "e")
	assert.Len(t, specs, 1)

	// Leading comma
	specs = parseOrderBy(",e.Name ASC", "e")
	assert.Len(t, specs, 1)
}

// ============================================================================
// Bug Fix: filterItems pre-allocation
// ============================================================================

func TestFilterItemsPreAllocation(t *testing.T) {
	cmd := &EachCommand{
		Items:  "items",
		Var:    "e",
		Select: "e > 5",
	}
	ctx := NewContext(nil)
	items := []any{1, 2, 3, 10, 20, 30}
	filtered, err := cmd.filterItems(items, ctx)
	require.NoError(t, err)
	assert.Len(t, filtered, 3)
	assert.Equal(t, 10, filtered[0])
}

// ============================================================================
// Integration: Full pipeline with all enhancements
// ============================================================================

func TestFullPipelineWithEnhancements(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	var debugBuf bytes.Buffer
	var progressCalls int

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
		},
	}

	outPath := filepath.Join(testdataDir(t), "full_pipeline_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, data,
		WithDebugWriter(&debugBuf),
		WithProgressFunc(func(p FillProgress) { progressCalls++ }),
		WithRecalculateOnOpen(true),
	)
	require.NoError(t, err)

	// Verify debug output
	trace := debugBuf.String()
	assert.Contains(t, trace, "[area]")
	assert.Contains(t, trace, "[done]")

	// Verify progress was reported
	assert.True(t, progressCalls > 0)

	// Verify output is valid
	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	v1, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v1)
	v3, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "Carol", v3)
}
