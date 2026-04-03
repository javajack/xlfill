package xlfill

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================================
// Edge case / bug fix tests
// ============================================================================

// Bug: getField should handle nil and non-struct/map gracefully
func TestGetFieldEdgeCases(t *testing.T) {
	assert.Nil(t, getField(nil, "Name"))
	assert.Nil(t, getField("string", "Name"))
	assert.Nil(t, getField(42, "Name"))

	// Map field
	m := map[string]any{"Name": "Alice"}
	assert.Equal(t, "Alice", getField(m, "Name"))
	assert.Nil(t, getField(m, "Missing"))

	// Struct field
	type Person struct{ Name string }
	p := Person{Name: "Bob"}
	assert.Equal(t, "Bob", getField(p, "Name"))
	assert.Nil(t, getField(p, "Missing"))

	// Pointer to struct
	assert.Equal(t, "Bob", getField(&p, "Name"))
}

// Bug: compareValues with mixed types should not panic
func TestCompareValuesEdgeCases(t *testing.T) {
	assert.Equal(t, 0, compareValues(nil, nil))
	assert.Equal(t, -1, compareValues(nil, "a"))
	assert.Equal(t, 1, compareValues("a", nil))
	assert.Equal(t, 0, compareValues("abc", "abc"))
	assert.Equal(t, -1, compareValues("abc", "xyz"))
	assert.Equal(t, 0, compareValues(42, 42))
	assert.Equal(t, -1, compareValues(1, 2))
	assert.Equal(t, -1, compareValues(1, "a")) // numeric vs string
}

// Bug: toSlice edge cases
func TestToSliceEdgeCases(t *testing.T) {
	result, err := toSlice(nil)
	assert.NoError(t, err)
	assert.Nil(t, result)

	_, err = toSlice("string")
	assert.Error(t, err)

	_, err = toSlice(42)
	assert.Error(t, err)

	// Array (not slice)
	arr := [3]int{1, 2, 3}
	result, err = toSlice(arr)
	assert.NoError(t, err)
	assert.Len(t, result, 3)
}

// Bug: SafeSheetName edge cases
func TestSafeSheetNameEdgeCases(t *testing.T) {
	// Exactly 31 chars (boundary)
	name31 := "ABCDEFGHIJKLMNOPQRSTUVWXYZ12345"
	assert.Equal(t, 31, len(name31))
	assert.Equal(t, name31, SafeSheetName(name31))

	// 32 chars — should truncate
	name32 := "ABCDEFGHIJKLMNOPQRSTUVWXYZ123456"
	assert.Equal(t, 31, len(SafeSheetName(name32)))

	// All forbidden chars
	assert.Equal(t, "_______", SafeSheetName("/\\:*?[]"))

	// Empty string
	assert.Equal(t, "", SafeSheetName(""))
}

// Bug: Context EvaluateCellValue with unbalanced delimiters
func TestEvaluateCellValueEdgeCases(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 5})

	// Normal expression
	val, ct, err := ctx.EvaluateCellValue("${x}")
	assert.NoError(t, err)
	assert.Equal(t, 5, val)
	assert.Equal(t, CellNumber, ct)

	// Mixed content
	val, ct, err = ctx.EvaluateCellValue("Value: ${x}")
	assert.NoError(t, err)
	assert.Equal(t, "Value: 5", val)
	assert.Equal(t, CellString, ct)

	// No expressions at all
	val, ct, err = ctx.EvaluateCellValue("plain text")
	assert.NoError(t, err)
	assert.Equal(t, "plain text", val)
	assert.Equal(t, CellString, ct)

	// Unbalanced — should return as plain text (no expressions found)
	val, ct, err = ctx.EvaluateCellValue("${broken")
	assert.NoError(t, err)
	assert.Equal(t, "${broken", val)
	assert.Equal(t, CellString, ct)
}

// Bug: inferCellType should handle all types correctly
func TestInferCellType(t *testing.T) {
	assert.Equal(t, CellBlank, inferCellType(nil))
	assert.Equal(t, CellBoolean, inferCellType(true))
	assert.Equal(t, CellNumber, inferCellType(42))
	assert.Equal(t, CellNumber, inferCellType(3.14))
	assert.Equal(t, CellNumber, inferCellType(int64(100)))
	assert.Equal(t, CellNumber, inferCellType(uint(10)))
	assert.Equal(t, CellString, inferCellType("hello"))
	assert.Equal(t, CellString, inferCellType(struct{}{})) // unknown type
}

// Bug: RunVar double close should not panic
func TestRunVarDoubleClose(t *testing.T) {
	ctx := NewContext(map[string]any{})

	rv := NewRunVar(ctx, "x")
	rv.Set("hello")
	assert.Equal(t, "hello", ctx.GetVar("x"))

	rv.Close()
	assert.Nil(t, ctx.GetVar("x"))

	// Double close should not panic
	assert.NotPanics(t, func() { rv.Close() })
}

// Bug: RunVar with index double close
func TestRunVarWithIndexDoubleClose(t *testing.T) {
	ctx := NewContext(map[string]any{})

	rv := NewRunVarWithIndex(ctx, "item", "idx")
	rv.SetWithIndex("value", 5)
	assert.Equal(t, "value", ctx.GetVar("item"))
	assert.Equal(t, 5, ctx.GetVar("idx"))

	rv.Close()
	assert.Nil(t, ctx.GetVar("item"))
	assert.Nil(t, ctx.GetVar("idx"))

	// Double close
	assert.NotPanics(t, func() { rv.Close() })
}

// Test nested RunVar save/restore
func TestRunVarNestedSaveRestore(t *testing.T) {
	ctx := NewContext(map[string]any{})

	// Outer loop sets "e"
	rv1 := NewRunVar(ctx, "e")
	rv1.Set("outer")
	assert.Equal(t, "outer", ctx.GetVar("e"))

	// Inner loop sets "e" (saves "outer")
	rv2 := NewRunVar(ctx, "e")
	rv2.Set("inner")
	assert.Equal(t, "inner", ctx.GetVar("e"))

	// Inner close restores "outer"
	rv2.Close()
	assert.Equal(t, "outer", ctx.GetVar("e"))

	// Outer close removes
	rv1.Close()
	assert.Nil(t, ctx.GetVar("e"))
}

// Test CellData Reset
func TestCellDataReset(t *testing.T) {
	cd := &CellData{
		Ref:             NewCellRef("Sheet1", 0, 0),
		TargetPositions: []CellRef{{Row: 1}, {Row: 2}},
		EvalResult:      "hello",
	}
	cd.Reset()
	assert.Len(t, cd.TargetPositions, 0)
	assert.Nil(t, cd.EvalResult)
}

// Test ExpressionSegment parsing edge cases
func TestParseExpressionsEdgeCases(t *testing.T) {
	// Nested braces
	segs := ParseExpressions("${a > 0 ? b : c}", "${", "}")
	require.Len(t, segs, 1)
	assert.True(t, segs[0].IsExpression)

	// Multiple expressions
	segs = ParseExpressions("${a} and ${b}", "${", "}")
	require.Len(t, segs, 3)
	assert.True(t, segs[0].IsExpression)
	assert.False(t, segs[1].IsExpression)
	assert.True(t, segs[2].IsExpression)

	// No expressions
	segs = ParseExpressions("plain", "${", "}")
	require.Len(t, segs, 1)
	assert.False(t, segs[0].IsExpression)
}

// Test ColToName/NameToCol round-trip
func TestColNameRoundTrip(t *testing.T) {
	for i := 0; i < 100; i++ {
		name := ColToName(i)
		col, err := NameToCol(name)
		assert.NoError(t, err)
		assert.Equal(t, i, col, "round-trip failed for col %d -> %s -> %d", i, name, col)
	}

	// Edge cases
	assert.Equal(t, "A", ColToName(0))
	assert.Equal(t, "Z", ColToName(25))
	assert.Equal(t, "AA", ColToName(26))
	assert.Equal(t, "AZ", ColToName(51))
	assert.Equal(t, "BA", ColToName(52))
}

// Test CellType String
func TestCellTypeString(t *testing.T) {
	assert.Equal(t, "Blank", CellBlank.String())
	assert.Equal(t, "String", CellString.String())
	assert.Equal(t, "Number", CellNumber.String())
	assert.Equal(t, "Boolean", CellBoolean.String())
	assert.Equal(t, "Date", CellDate.String())
	assert.Equal(t, "Formula", CellFormula.String())
	assert.Equal(t, "Error", CellError.String())
	assert.Equal(t, "Unknown", CellType(99).String())
}

// Test AreaRef operations
func TestAreaRefOperations(t *testing.T) {
	ref, err := ParseAreaRef("Sheet1!A1:C5")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ref.SheetName())

	size := ref.Size()
	assert.Equal(t, 3, size.Width)
	assert.Equal(t, 5, size.Height)

	assert.True(t, ref.Contains(NewCellRef("Sheet1", 2, 1)))
	assert.False(t, ref.Contains(NewCellRef("Sheet1", 10, 0)))
	assert.False(t, ref.Contains(NewCellRef("Sheet2", 0, 0)))
}

// Test Size operations
func TestSizeOperations(t *testing.T) {
	s1 := Size{Width: 3, Height: 5}
	s2 := Size{Width: 2, Height: 1}

	added := s1.Add(s2)
	assert.Equal(t, Size{Width: 5, Height: 6}, added)

	minus := s1.Minus(s2)
	assert.Equal(t, Size{Width: 1, Height: 4}, minus)

	assert.Equal(t, "(3x5)", s1.String())
}

// Test DebugTracer indent/dedent
func TestDebugTracerIndent(t *testing.T) {
	var buf bytes.Buffer
	d := NewDebugTracer(&buf)

	assert.Equal(t, 0, d.indent)
	d.Indent()
	assert.Equal(t, 1, d.indent)
	d.Indent()
	assert.Equal(t, 2, d.indent)
	d.Dedent()
	assert.Equal(t, 1, d.indent)
	d.Dedent()
	assert.Equal(t, 0, d.indent)
	d.Dedent() // should not go below 0
	assert.Equal(t, 0, d.indent)
}

// Test compiled template with options
func TestCompiledTemplateWithOptions(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl, WithRecalculateOnOpen(true))
	require.NoError(t, err)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	b, err := compiled.FillBytes(data)
	require.NoError(t, err)
	assert.True(t, len(b) > 0)
}

// Test compiled template FillWriter
func TestCompiledTemplateFillWriter(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl)
	require.NoError(t, err)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	var buf bytes.Buffer
	err = compiled.FillWriter(data, &buf)
	require.NoError(t, err)
	assert.True(t, buf.Len() > 0)
}

// Test toInt edge cases
func TestToIntEdgeCases(t *testing.T) {
	assert.Equal(t, 5, toInt(5))
	assert.Equal(t, 5, toInt(int64(5)))
	assert.Equal(t, 5, toInt(float64(5.7))) // truncates
	assert.Equal(t, 5, toInt(float32(5.2))) // truncates
	assert.Equal(t, 5, toInt("5"))
	assert.Equal(t, 1, toInt("not a number")) // default
	assert.Equal(t, 1, toInt(nil))            // default
}

// Test GroupData struct
func TestGroupDataStruct(t *testing.T) {
	gd := GroupData{
		Item:  map[string]any{"Name": "Engineering"},
		Items: []any{map[string]any{"Name": "Alice"}, map[string]any{"Name": "Bob"}},
	}
	assert.Len(t, gd.Items, 2)
}

// Test each command with empty items after filter
func TestEachCommandEmptyAfterFilter(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:each(items="items" var="e" select="e > 100" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "empty_filter_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "empty_filter_out.xlsx")
	defer os.Remove(outPath)

	// All items are < 100, so filter removes everything
	err := Fill(path, outPath, map[string]any{
		"items": []int{1, 2, 3},
	})
	require.NoError(t, err)
}

// Test Filler warnings are reset between calls
func TestFillerWarningsReset(t *testing.T) {
	tmpl := createTemplateWithUnknownCommand(t)
	defer os.Remove(tmpl)

	filler := NewFiller(WithTemplate(tmpl))

	tx1, err := filler.openTemplate()
	require.NoError(t, err)
	_, err = filler.BuildAreas(tx1)
	require.NoError(t, err)
	tx1.Close()

	assert.Len(t, filler.Warnings(), 1)

	// Second call should reset warnings
	tx2, err := filler.openTemplate()
	require.NoError(t, err)
	_, err = filler.BuildAreas(tx2)
	require.NoError(t, err)
	tx2.Close()

	assert.Len(t, filler.Warnings(), 1) // fresh warnings, not accumulated
}

// Test extractTopLevelVars with dot-access chains
func TestExtractTopLevelVarsDotAccess(t *testing.T) {
	vars := extractTopLevelVars("a.b.c + d.e")
	assert.Contains(t, vars, "a")
	assert.Contains(t, vars, "d")
	assert.NotContains(t, vars, "b") // b is after dot
	assert.NotContains(t, vars, "c")
	assert.NotContains(t, vars, "e")
}

// Bug fix: extractTopLevelVars should skip string literals
func TestExtractTopLevelVarsSkipsStringLiterals(t *testing.T) {
	// Single-quoted strings
	vars := extractTopLevelVars("'hello' + name")
	assert.Contains(t, vars, "name")
	assert.NotContains(t, vars, "hello") // inside string literal

	// Double-quoted strings
	vars = extractTopLevelVars(`"world" + variable`)
	assert.Contains(t, vars, "variable")
	assert.NotContains(t, vars, "world")

	// Mixed quotes with escaped chars
	vars = extractTopLevelVars(`"it\'s" + x`)
	assert.Contains(t, vars, "x")

	// String with dot-access inside (should not extract)
	vars = extractTopLevelVars(`'e.Name' + actual`)
	assert.Contains(t, vars, "actual")
	assert.NotContains(t, vars, "e")
	assert.NotContains(t, vars, "Name")

	// Backtick strings
	vars = extractTopLevelVars("`template` + data")
	assert.Contains(t, vars, "data")
	assert.NotContains(t, vars, "template")
}

// Bug fix: KnownNames should be sorted
func TestKnownNamesSorted(t *testing.T) {
	reg := NewCommandRegistry()
	names := reg.KnownNames()

	for i := 1; i < len(names); i++ {
		assert.True(t, names[i-1] <= names[i],
			"KnownNames not sorted: %q > %q", names[i-1], names[i])
	}
}

// Bug fix: RepeatCommand nil area check comes before count eval
func TestRepeatNilAreaErrorBeforeCountEval(t *testing.T) {
	cmd := &RepeatCommand{Count: "invalid_expr", Direction: "DOWN"}
	ctx := NewContext(nil)

	// Should fail with "no area" not "evaluate count"
	_, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, nil)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "no area")
}

// Bug fix: hyperlink builtin preserved through differential updates
func TestHyperlinkBuiltinPreservedInDifferentialUpdate(t *testing.T) {
	ctx := NewContext(map[string]any{})

	// Initial build should have hyperlink
	m := ctx.ToMap()
	assert.NotNil(t, m["hyperlink"])

	// Set and remove a runVar — hyperlink should survive
	ctx.setRunVar("x", 1)
	ctx.removeRunVar("x")
	m = ctx.ToMap()
	assert.NotNil(t, m["hyperlink"], "hyperlink builtin should survive differential updates")

	// Even if someone names a runVar "hyperlink" and removes it
	ctx.setRunVar("hyperlink", "override")
	m = ctx.ToMap()
	assert.Equal(t, "override", m["hyperlink"])

	ctx.removeRunVar("hyperlink")
	m = ctx.ToMap()
	// Should restore the builtin, not delete
	assert.NotNil(t, m["hyperlink"], "hyperlink builtin should be restored after runVar removal")
}

// Bug fix: compiled template copies all options
func TestCompiledTemplateCopiesAllOptions(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl,
		WithKeepTemplateSheet(true),
		WithHideTemplateSheet(true),
		WithStrictMode(true),
	)
	require.NoError(t, err)

	// Verify options propagated
	assert.True(t, compiled.opts.keepTemplateSheet)
	assert.True(t, compiled.opts.hideTemplateSheet)
	assert.True(t, compiled.opts.strictMode)
}

// Bug fix: DebugTracer thread safety
func TestDebugTracerConcurrentSafety(t *testing.T) {
	var buf bytes.Buffer
	tracer := NewDebugTracer(&buf)

	// Run concurrent Indent/Dedent/BeforeTransformCell
	done := make(chan struct{})
	for i := 0; i < 10; i++ {
		go func() {
			defer func() { done <- struct{}{} }()
			for j := 0; j < 100; j++ {
				tracer.Indent()
				tracer.BeforeTransformCell(
					NewCellRef("Sheet1", 0, 0),
					NewCellRef("Sheet1", 0, 0),
					NewContext(nil), nil)
				tracer.Dedent()
			}
		}()
	}
	for i := 0; i < 10; i++ {
		<-done
	}

	assert.Equal(t, int64(1000), tracer.cellCount.Load())
}

// Bug fix: additional expr keywords
func TestAdditionalExprKeywords(t *testing.T) {
	assert.True(t, isExprKeyword("range"))
	assert.True(t, isExprKeyword("reverse"))
	assert.True(t, isExprKeyword("compact"))
	assert.True(t, isExprKeyword("unique"))
	assert.True(t, isExprKeyword("keys"))
	assert.True(t, isExprKeyword("values"))
	assert.True(t, isExprKeyword("abs"))
	assert.True(t, isExprKeyword("ceil"))
	assert.True(t, isExprKeyword("floor"))
	assert.True(t, isExprKeyword("round"))
	assert.True(t, isExprKeyword("now"))
	assert.True(t, isExprKeyword("type"))
	assert.True(t, isExprKeyword("reduce"))
}

// Test StyleOverride fields
func TestStyleOverrideFields(t *testing.T) {
	bold := true
	italic := false
	color := "#FF0000"
	fill := "#00FF00"
	size := 14.0

	s := &StyleOverride{
		Bold:      &bold,
		Italic:    &italic,
		FontColor: &color,
		FillColor: &fill,
		FontSize:  &size,
	}

	assert.True(t, *s.Bold)
	assert.False(t, *s.Italic)
	assert.Equal(t, "#FF0000", *s.FontColor)
	assert.Equal(t, "#00FF00", *s.FillColor)
	assert.Equal(t, 14.0, *s.FontSize)
}

// Test HyperlinkValue String method
func TestHyperlinkValueString(t *testing.T) {
	h1 := HyperlinkValue{URL: "https://example.com", Display: "Example"}
	assert.Equal(t, "Example", h1.String())

	h2 := HyperlinkValue{URL: "https://example.com"}
	assert.Equal(t, "https://example.com", h2.String())
}

// Test Hyperlink function
func TestHyperlinkFunction(t *testing.T) {
	h := Hyperlink("https://example.com", "Example")
	assert.Equal(t, "https://example.com", h.URL)
	assert.Equal(t, "Example", h.Display)
}

// Test isIdentStart/isIdentPart
func TestIdentCharHelpers(t *testing.T) {
	assert.True(t, isIdentStart('a'))
	assert.True(t, isIdentStart('Z'))
	assert.True(t, isIdentStart('_'))
	assert.False(t, isIdentStart('0'))
	assert.False(t, isIdentStart('.'))

	assert.True(t, isIdentPart('a'))
	assert.True(t, isIdentPart('0'))
	assert.True(t, isIdentPart('_'))
	assert.False(t, isIdentPart('.'))
	assert.False(t, isIdentPart(' '))
}
