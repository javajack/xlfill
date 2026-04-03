package xlfill

import (
	"bytes"
	"os"
	"path/filepath"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================================
// Streaming Mode Tests
// ============================================================================

func TestStreamingBasic(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "streaming_basic_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
		},
	}

	err := Fill(tmpl, outPath, data, WithStreaming(true))
	require.NoError(t, err)

	// Verify output
	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	// Row 1 is header, data starts at row 2
	v1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "Name", v1) // header preserved

	v2, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v2)

	v3, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Bob", v3)

	v4, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "Carol", v4)
}

func TestStreamingLargeDataset(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "streaming_large_out.xlsx")
	defer os.Remove(outPath)

	// Create 1000 employees
	employees := make([]map[string]any, 1000)
	for i := range employees {
		employees[i] = map[string]any{
			"Name":   "Employee" + string(rune('A'+i%26)),
			"Age":    20 + i%50,
			"Salary": 1000 + i*10,
		}
	}

	err := Fill(tmpl, outPath, map[string]any{"employees": employees}, WithStreaming(true))
	require.NoError(t, err)

	// Verify first and last rows
	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	v1, _ := f.GetCellValue("Sheet1", "A1")
	assert.NotEmpty(t, v1)

	v1000, _ := f.GetCellValue("Sheet1", "A1000")
	assert.NotEmpty(t, v1000)
}

func TestStreamingFillBytes(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	b, err := FillBytes(tmpl, data, WithStreaming(true))
	require.NoError(t, err)
	assert.True(t, len(b) > 0)

	// Verify content
	f, err := excelize.OpenReader(bytes.NewReader(b))
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v)
}

func TestStreamingTransformerImplementsInterface(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	stx, err := NewStreamingTransformer(etx, "Sheet1")
	require.NoError(t, err)

	// Verify interface satisfaction
	var _ Transformer = stx
	var _ CellReader = stx
	var _ CellWriter = stx
	var _ SheetManager = stx
}

func TestStreamingImageNotSupported(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	stx, err := NewStreamingTransformer(etx, "Sheet1")
	require.NoError(t, err)

	err = stx.AddImage("Sheet1", "A1", []byte{}, "PNG", 1.0, 1.0)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "not supported in streaming mode")
}

func TestStreamingFlush(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	stx, err := NewStreamingTransformer(etx, "Sheet1")
	require.NoError(t, err)

	// Buffer some cells
	require.NoError(t, stx.SetCellValue(NewCellRef("Sheet1", 0, 0), "A1"))
	require.NoError(t, stx.SetCellValue(NewCellRef("Sheet1", 0, 1), "B1"))
	require.NoError(t, stx.SetCellValue(NewCellRef("Sheet1", 1, 0), "A2"))

	// Flush
	require.NoError(t, stx.Flush())

	// Double flush should be safe
	require.NoError(t, stx.Flush())
}

// ============================================================================
// Parallel Processing Tests
// ============================================================================

func TestParallelBasic(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_basic_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
			{"Name": "Dave", "Age": 28, "Salary": 1500},
		},
	}

	err := Fill(tmpl, outPath, data, WithParallelism(2))
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	// Verify all rows are present and in correct order
	v1, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v1)

	v2, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Bob", v2)

	v3, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "Carol", v3)

	v4, _ := f.GetCellValue("Sheet1", "A5")
	assert.Equal(t, "Dave", v4)
}

func TestParallelLargeDataset(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_large_out.xlsx")
	defer os.Remove(outPath)

	employees := make([]map[string]any, 100)
	for i := range employees {
		employees[i] = map[string]any{
			"Name":   "E" + string(rune('A'+i%26)),
			"Age":    20 + i,
			"Salary": 1000 + i*100,
		}
	}

	err := Fill(tmpl, outPath, map[string]any{"employees": employees}, WithParallelism(4))
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	// Spot check
	v1, _ := f.GetCellValue("Sheet1", "A2")
	assert.NotEmpty(t, v1)

	v100, _ := f.GetCellValue("Sheet1", "A101")
	assert.NotEmpty(t, v100)
}

func TestParallelWithProgressReporting(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_progress_out.xlsx")
	defer os.Remove(outPath)

	var mu sync.Mutex
	var progressCalls int

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
			{"Name": "Dave", "Age": 28, "Salary": 1500},
		},
	}

	err := Fill(tmpl, outPath, data,
		WithParallelism(2),
		WithProgressFunc(func(p FillProgress) {
			mu.Lock()
			progressCalls++
			mu.Unlock()
		}),
	)
	require.NoError(t, err)
	assert.True(t, progressCalls > 0)
}

func TestParallelDisabledForSmallDatasets(t *testing.T) {
	// With parallelism=4 but only 2 items, should fall back to sequential
	// (because len(items) < parallelism)
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_small_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
		},
	}

	err := Fill(tmpl, outPath, data, WithParallelism(4))
	require.NoError(t, err)
}

func TestParallelRaceDetection(t *testing.T) {
	// This test is specifically designed to catch race conditions when
	// run with -race. It uses enough parallelism and data to trigger races.
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_race_out.xlsx")
	defer os.Remove(outPath)

	employees := make([]map[string]any, 50)
	for i := range employees {
		employees[i] = map[string]any{
			"Name":   "Employee",
			"Age":    i,
			"Salary": i * 100,
		}
	}

	err := Fill(tmpl, outPath, map[string]any{"employees": employees}, WithParallelism(8))
	require.NoError(t, err)
}

func TestParallelFallbackForVariableHeight(t *testing.T) {
	// Create a template with nested each (variable height) — should fall back
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

	path := filepath.Join(testdataDir(t), "parallel_nested_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "parallel_nested_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"departments": []map[string]any{
			{
				"Name":      "Engineering",
				"Employees": []map[string]any{{"Name": "Alice"}, {"Name": "Bob"}},
			},
			{
				"Name":      "Marketing",
				"Employees": []map[string]any{{"Name": "Carol"}},
			},
		},
	}

	// Should work even with parallel — falls back to sequential for nested each
	err := Fill(path, outPath, data, WithParallelism(4))
	require.NoError(t, err)
}

// ============================================================================
// Concurrent Transformer Tests
// ============================================================================

func TestConcurrentTransformerInterface(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	ct := NewConcurrentTransformer(etx)

	var _ Transformer = ct
	var _ CellReader = ct
	var _ CellWriter = ct
	var _ SheetManager = ct

	assert.Equal(t, etx, ct.Inner())
}

func TestConcurrentTransformerReadOperations(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	ct := NewConcurrentTransformer(etx)

	// Read operations should work
	sheets := ct.GetSheetNames()
	assert.Contains(t, sheets, "Sheet1")

	commented := ct.GetCommentedCells()
	assert.True(t, len(commented) > 0)
}

// ============================================================================
// Context Clone Tests
// ============================================================================

func TestContextClone(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 1, "y": 2})
	ctx.setRunVar("loop", "outer")

	clone := ctx.Clone()

	// Clone should share data
	assert.Equal(t, 1, clone.GetVar("x"))
	assert.Equal(t, 2, clone.GetVar("y"))

	// Clone should NOT have runVars from original
	assert.Nil(t, clone.GetVar("loop"))

	// Modifying clone's runVars should not affect original
	clone.setRunVar("loop", "inner")
	assert.Equal(t, "outer", ctx.GetVar("loop"))
	assert.Equal(t, "inner", clone.GetVar("loop"))
}

func TestContextCloneIndependentEvaluation(t *testing.T) {
	ctx := NewContext(map[string]any{"base": 10})

	clone1 := ctx.Clone()
	clone2 := ctx.Clone()

	clone1.setRunVar("i", 1)
	clone2.setRunVar("i", 2)

	v1, err := clone1.Evaluate("base + i")
	require.NoError(t, err)
	assert.Equal(t, 11, v1)

	v2, err := clone2.Evaluate("base + i")
	require.NoError(t, err)
	assert.Equal(t, 12, v2)
}

// ============================================================================
// isFixedHeight Tests
// ============================================================================

func TestIsFixedHeight(t *testing.T) {
	// Area with no bindings → fixed height
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{3, 2}, nil)
	assert.True(t, isAreaFixedHeight(area))

	// Area with nested each → NOT fixed height
	eachCmd := &EachCommand{Items: "items", Var: "e", Direction: "DOWN"}
	area.AddCommand(eachCmd, NewCellRef("Sheet1", 1, 0), Size{3, 1})
	assert.False(t, isAreaFixedHeight(area))
}

func TestIsFixedHeightWithRepeat(t *testing.T) {
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{3, 2}, nil)
	repeatCmd := &RepeatCommand{Count: "5", Var: "i", Direction: "DOWN"}
	area.AddCommand(repeatCmd, NewCellRef("Sheet1", 1, 0), Size{3, 1})
	assert.False(t, isAreaFixedHeight(area))
}

func TestIsFixedHeightWithIf(t *testing.T) {
	// If with same-sized branches → fixed height
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{3, 2}, nil)
	ifCmd := &IfCommand{
		Condition: "true",
		IfArea:    NewArea(NewCellRef("Sheet1", 1, 0), Size{3, 1}, nil),
		ElseArea:  NewArea(NewCellRef("Sheet1", 1, 0), Size{3, 1}, nil),
	}
	area.AddCommand(ifCmd, NewCellRef("Sheet1", 1, 0), Size{3, 1})
	assert.True(t, isAreaFixedHeight(area))

	// If with different-sized branches → NOT fixed height
	ifCmd2 := &IfCommand{
		Condition: "true",
		IfArea:    NewArea(NewCellRef("Sheet1", 1, 0), Size{3, 1}, nil),
		ElseArea:  NewArea(NewCellRef("Sheet1", 1, 0), Size{3, 2}, nil),
	}
	area2 := NewArea(NewCellRef("Sheet1", 0, 0), Size{3, 3}, nil)
	area2.AddCommand(ifCmd2, NewCellRef("Sheet1", 1, 0), Size{3, 2})
	assert.False(t, isAreaFixedHeight(area2))
}

// ============================================================================
// Streaming + Parallel Mutual Exclusion
// ============================================================================

func TestStreamingAndParallelMutualExclusion(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "streaming_parallel_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
			{"Name": "Dave", "Age": 28, "Salary": 1500},
		},
	}

	// When both streaming and parallel are set, parallel wins
	err := Fill(tmpl, outPath, data, WithStreaming(true), WithParallelism(2))
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	v1, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v1)
}

// ============================================================================
// Edge cases
// ============================================================================

func TestParallelism1IsSequential(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_1_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	// Parallelism of 1 should work (sequential)
	err := Fill(tmpl, outPath, data, WithParallelism(1))
	require.NoError(t, err)
}

func TestStreamingDisabledByDefault(t *testing.T) {
	opts := defaultOptions()
	assert.False(t, opts.streaming)
	assert.Equal(t, 0, opts.parallelism)
}

// ============================================================================
// Benchmarks
// ============================================================================

func BenchmarkFill_1000Rows_Parallel2(b *testing.B) {
	tmpl := createBenchTemplate2(b)
	defer os.Remove(tmpl)

	data := createBenchData2(1000)
	outPath := filepath.Join("testdata", "bench_parallel_out.xlsx")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		Fill(tmpl, outPath, data, WithParallelism(2))
	}
	b.StopTimer()
	os.Remove(outPath)
}

func BenchmarkFill_1000Rows_Parallel4(b *testing.B) {
	tmpl := createBenchTemplate2(b)
	defer os.Remove(tmpl)

	data := createBenchData2(1000)
	outPath := filepath.Join("testdata", "bench_parallel4_out.xlsx")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		Fill(tmpl, outPath, data, WithParallelism(4))
	}
	b.StopTimer()
	os.Remove(outPath)
}

func BenchmarkFill_1000Rows_Streaming(b *testing.B) {
	tmpl := createBenchTemplate2(b)
	defer os.Remove(tmpl)

	data := createBenchData2(1000)
	outPath := filepath.Join("testdata", "bench_streaming_out.xlsx")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		Fill(tmpl, outPath, data, WithStreaming(true))
	}
	b.StopTimer()
	os.Remove(outPath)
}

func createBenchTemplate2(tb testing.TB) string {
	tb.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Age")
	f.SetCellValue(sheet, "C1", "Salary")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Age}")
	f.SetCellValue(sheet, "C2", "${e.Salary}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="C2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: `jx:each(items="employees" var="e" lastCell="C2")`,
	})

	path := filepath.Join("testdata", "bench_template.xlsx")
	os.MkdirAll("testdata", 0o755)
	f.SaveAs(path)
	return path
}

func createBenchData2(n int) map[string]any {
	employees := make([]map[string]any, n)
	for i := range employees {
		employees[i] = map[string]any{
			"Name":   "Employee",
			"Age":    20 + i%50,
			"Salary": 1000 + i*10,
		}
	}
	return map[string]any{"employees": employees}
}
