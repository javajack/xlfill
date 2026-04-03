package xlfill

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================================
// SuggestMode Tests
// ============================================================================

func TestSuggestModeSequentialForSmallData(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, map[string]any{"itemCount": 10})
	require.NoError(t, err)
	assert.Equal(t, ModeSequential, s.Mode)
	assert.True(t, len(s.Reasons) > 0)
}

func TestSuggestModeStreamingForLargeData(t *testing.T) {
	// Basic template has no formulas, images, or hyperlinks → streaming eligible
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, map[string]any{"itemCount": 50000})
	require.NoError(t, err)
	assert.Equal(t, ModeStreaming, s.Mode)
	assert.Contains(t, s.Reasons[0], "large dataset")
}

func TestSuggestModeParallelForModerateData(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, map[string]any{"itemCount": 500})
	require.NoError(t, err)
	// On multi-core machines → parallel; on single-core → streaming fallback at >=1000
	// Either is acceptable; verify the suggestion is valid regardless
	switch s.Mode {
	case ModeParallel:
		assert.True(t, s.Parallelism > 0)
		assert.True(t, s.Parallelism <= 8)
	case ModeSequential:
		// Single-core machine — parallel not available, 500 < 1000 streaming threshold
	default:
		t.Fatalf("unexpected mode %v for 500 items", s.Mode)
	}
}

func TestSuggestModeBlocksStreamingForFormulas(t *testing.T) {
	tmpl := createFormulaTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, map[string]any{"itemCount": 50000})
	require.NoError(t, err)
	// Should NOT suggest streaming because of formulas
	assert.NotEqual(t, ModeStreaming, s.Mode)
}

func TestSuggestModeBlocksStreamingForImages(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")
	f.SetCellValue(sheet, "B1", "img")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")
jx:each(items="employees" var="e" lastCell="B1")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "B1", Author: "xlfill",
		Text: `jx:image(src="e.Photo" lastCell="B1")`,
	})

	path := filepath.Join(testdataDir(t), "suggest_image_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	s, err := SuggestMode(path, map[string]any{"itemCount": 50000})
	require.NoError(t, err)
	assert.NotEqual(t, ModeStreaming, s.Mode)
}

func TestSuggestModeBlocksParallelForNestedEach(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${dept.Name}")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")
jx:each(items="departments" var="dept" lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="dept.Employees" var="e" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "suggest_nested_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	s, err := SuggestMode(path, map[string]any{"itemCount": 500})
	require.NoError(t, err)
	// Nested each means variable height → no parallel
	assert.NotEqual(t, ModeParallel, s.Mode)
}

func TestSuggestModeBlocksStreamingForHyperlinks(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${hyperlink(e.URL, e.Name)}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:each(items="items" var="e" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "suggest_hyperlink_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	s, err := SuggestMode(path, map[string]any{"itemCount": 50000})
	require.NoError(t, err)
	assert.NotEqual(t, ModeStreaming, s.Mode)
}

func TestSuggestModeNilHint(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, nil)
	require.NoError(t, err)
	// With no item count hint, should default to sequential
	assert.Equal(t, ModeSequential, s.Mode)
}

func TestSuggestModeBlocksParallelForDirectionRight(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:each(items="items" var="e" direction="RIGHT" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "suggest_right_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	s, err := SuggestMode(path, map[string]any{"itemCount": 500})
	require.NoError(t, err)
	assert.NotEqual(t, ModeParallel, s.Mode)
	assert.NotEqual(t, ModeStreaming, s.Mode)
}

// ============================================================================
// WithAutoMode Integration Tests
// ============================================================================

func TestWithAutoModeSmallData(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "automode_small_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	err := Fill(tmpl, outPath, data, WithAutoMode(map[string]any{"itemCount": 1}))
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v)
}

func TestWithAutoModeStreamingHint(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "automode_streaming_out.xlsx")
	defer os.Remove(outPath)

	employees := make([]map[string]any, 100)
	for i := range employees {
		employees[i] = map[string]any{"Name": "E", "Age": i, "Salary": i * 100}
	}

	// Hint says 50K items → should auto-select streaming
	err := Fill(tmpl, outPath, map[string]any{"employees": employees},
		WithAutoMode(map[string]any{"itemCount": 50000}))
	require.NoError(t, err)
}

func TestWithAutoModeNoHint(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "automode_nohint_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	// No hint → auto-mode still works (defaults to sequential)
	err := Fill(tmpl, outPath, data, WithAutoMode(nil))
	require.NoError(t, err)
}

// ============================================================================
// ModeSuggestion.Apply Tests
// ============================================================================

func TestModeSuggestionApply(t *testing.T) {
	// Streaming
	s := &ModeSuggestion{Mode: ModeStreaming}
	opts := s.Apply()
	assert.Len(t, opts, 1)

	// Parallel
	s2 := &ModeSuggestion{Mode: ModeParallel, Parallelism: 4}
	opts2 := s2.Apply()
	assert.Len(t, opts2, 1)

	// Sequential
	s3 := &ModeSuggestion{Mode: ModeSequential}
	opts3 := s3.Apply()
	assert.Len(t, opts3, 0) // no options added
}

// ============================================================================
// Mode String Tests
// ============================================================================

func TestModeString(t *testing.T) {
	assert.Equal(t, "sequential", ModeSequential.String())
	assert.Equal(t, "streaming", ModeStreaming.String())
	assert.Equal(t, "parallel", ModeParallel.String())
	assert.Equal(t, "unknown", Mode(99).String())
}

// ============================================================================
// Template Analysis Tests
// ============================================================================

func TestAnalyzeTemplateDetectsFormulas(t *testing.T) {
	tmpl := createFormulaTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	filler := NewFiller(WithTemplate(tmpl))
	areas, err := filler.BuildAreas(etx)
	require.NoError(t, err)

	a := analyzeTemplate(areas, etx)
	assert.True(t, a.hasFormulas)
}

func TestAnalyzeTemplateDetectsNestedEach(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${d.Name}")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A2")
jx:each(items="deps" var="d" lastCell="A2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="d.Emps" var="e" lastCell="A2")`,
	})

	path := filepath.Join(testdataDir(t), "analyze_nested_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	etx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer etx.Close()

	filler := NewFiller(WithTemplate(path))
	areas, err := filler.BuildAreas(etx)
	require.NoError(t, err)

	a := analyzeTemplate(areas, etx)
	assert.True(t, a.hasNestedEach)
	assert.False(t, a.isFixedHeight)
	assert.Equal(t, 2, a.eachCount)
	assert.Equal(t, 2, a.maxEachDepth)
}

func TestAnalyzeTemplateDetectsMergeCells(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")
jx:each(items="items" var="e" lastCell="B1")
jx:mergeCells(cols="2" lastCell="B1")`,
	})

	path := filepath.Join(testdataDir(t), "analyze_merge_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	etx, err := OpenTemplate(path)
	require.NoError(t, err)
	defer etx.Close()

	filler := NewFiller(WithTemplate(path))
	areas, err := filler.BuildAreas(etx)
	require.NoError(t, err)

	a := analyzeTemplate(areas, etx)
	assert.True(t, a.hasMergeCells)
}
