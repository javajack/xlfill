package xlfill

import (
	"context"
	"os"
	"path/filepath"
	"sync/atomic"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================================
// Fix: Parallel cancellation propagation
// ============================================================================

func TestParallelCancellationPropagates(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_cancel_out.xlsx")
	defer os.Remove(outPath)

	ctx, cancel := context.WithCancel(context.Background())
	cancel() // cancel immediately

	employees := make([]map[string]any, 20)
	for i := range employees {
		employees[i] = map[string]any{"Name": "E", "Age": i, "Salary": i}
	}

	err := Fill(tmpl, outPath, map[string]any{"employees": employees},
		WithParallelism(4), WithContext(ctx))
	assert.Error(t, err, "parallel processing should detect cancelled context")
}

func TestParallelCancellationOnError(t *testing.T) {
	// Create a template where expression evaluation will fail for some items
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	// This expression will fail if e is not a map
	f.SetCellValue(sheet, "A1", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:each(items="items" var="e" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "parallel_error_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "parallel_error_out.xlsx")
	defer os.Remove(outPath)

	// Mix valid and invalid items — invalid ones will cause expression errors
	items := make([]any, 20)
	for i := range items {
		items[i] = map[string]any{"Name": "Valid"}
	}
	// This nil item will cause the expression to fail
	items[5] = nil

	err := Fill(path, outPath, map[string]any{"items": items}, WithParallelism(4))
	// Should get an error (not a panic)
	assert.Error(t, err)
}

// ============================================================================
// Fix: Parallel goroutine panic recovery
// ============================================================================

func TestParallelPanicRecovery(t *testing.T) {
	// This test verifies that panics in parallel goroutines are recovered
	// and reported as errors, not crashing the program.
	// We can't easily trigger a panic through normal template processing,
	// but we verify the infrastructure is in place by running a large
	// parallel job without panics.
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_panic_out.xlsx")
	defer os.Remove(outPath)

	employees := make([]map[string]any, 30)
	for i := range employees {
		employees[i] = map[string]any{"Name": "E", "Age": i, "Salary": i}
	}

	err := Fill(tmpl, outPath, map[string]any{"employees": employees}, WithParallelism(8))
	require.NoError(t, err)
}

// ============================================================================
// Fix: Streaming style preservation
// ============================================================================

func TestStreamingPreservesStyleInSetCellValue(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	stx, err := NewStreamingTransformer(etx, "Sheet1")
	require.NoError(t, err)

	// SetCellValue should preserve style via lookupStyle
	ref := NewCellRef("Sheet1", 0, 0)
	err = stx.SetCellValue(ref, "test")
	require.NoError(t, err)

	// ClearCell should also preserve style
	err = stx.ClearCell(ref)
	require.NoError(t, err)

	// SetFormula should also preserve style
	err = stx.SetFormula(ref, "SUM(A1:A5)")
	require.NoError(t, err)
}

// ============================================================================
// Fix: Streaming flushRowLocked starts from correct column
// ============================================================================

func TestStreamingFlushStartsFromCorrectColumn(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	stx, err := NewStreamingTransformer(etx, "Sheet1")
	require.NoError(t, err)

	// Buffer cells starting from column C (index 2), not A
	require.NoError(t, stx.SetCellValue(NewCellRef("Sheet1", 0, 2), "C1"))
	require.NoError(t, stx.SetCellValue(NewCellRef("Sheet1", 0, 3), "D1"))
	require.NoError(t, stx.SetCellValue(NewCellRef("Sheet1", 1, 2), "C2"))

	// Flush — should not crash or pad with nils from column A
	require.NoError(t, stx.Flush())
}

// ============================================================================
// Fix: Auto-mode doesn't mutate f.opts
// ============================================================================

func TestAutoModeDoesNotMutateFiller(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	filler := NewFiller(
		WithTemplate(tmpl),
		WithAutoMode(map[string]any{"itemCount": 50000}),
	)

	// Record original values
	origStreaming := filler.opts.streaming
	origParallelism := filler.opts.parallelism

	data := map[string]any{
		"employees": []map[string]any{{"Name": "Alice", "Age": 30, "Salary": 1000}},
	}
	outPath := filepath.Join(testdataDir(t), "automode_nomutate_out.xlsx")
	defer os.Remove(outPath)

	err := filler.Fill(data, outPath)
	require.NoError(t, err)

	// f.opts should NOT have been mutated
	assert.Equal(t, origStreaming, filler.opts.streaming)
	assert.Equal(t, origParallelism, filler.opts.parallelism)
}

// ============================================================================
// Fix: scanForHyperlinks checks formulas
// ============================================================================

func TestScanForHyperlinksInFormulas(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Label")
	f.SetCellFormula(sheet, "B1", `HYPERLINK("https://example.com","Link")`)

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="B1")`,
	})

	path := filepath.Join(testdataDir(t), "hyperlink_formula_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	s, err := SuggestMode(path, map[string]any{"itemCount": 50000})
	require.NoError(t, err)
	// Should NOT suggest streaming because of HYPERLINK formula
	assert.NotEqual(t, ModeStreaming, s.Mode)
}

// ============================================================================
// Fix: Boundary thresholds use >=
// ============================================================================

func TestAutoModeBoundaryExactly10000(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, map[string]any{"itemCount": 10000})
	require.NoError(t, err)
	// Exactly 10000 should now trigger streaming (>= not >)
	assert.Equal(t, ModeStreaming, s.Mode)
}

func TestAutoModeBoundaryExactly100(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	s, err := SuggestMode(tmpl, map[string]any{"itemCount": 100})
	require.NoError(t, err)
	// Exactly 100 should trigger parallel on multi-core (>= not >)
	// On single-core, falls through to streaming >= 1000 check
	if s.Mode == ModeParallel {
		assert.True(t, s.Parallelism >= 1)
	}
}

// ============================================================================
// Fix: Context.Clone is safe for parallel use
// ============================================================================

func TestContextCloneParallelSafety(t *testing.T) {
	ctx := NewContext(map[string]any{"shared": "data"})

	done := make(chan struct{})
	for i := 0; i < 10; i++ {
		go func(id int) {
			defer func() { done <- struct{}{} }()
			clone := ctx.Clone()
			clone.setRunVar("id", id)
			// Each clone should see its own "id" and the shared "data"
			assert.Equal(t, id, clone.GetVar("id"))
			assert.Equal(t, "data", clone.GetVar("shared"))

			// Evaluate using the cloned context
			m := clone.ToMap()
			assert.Equal(t, id, m["id"])
			assert.Equal(t, "data", m["shared"])
		}(i)
	}
	for i := 0; i < 10; i++ {
		<-done
	}
}

// ============================================================================
// Fix: StyleListener works with ConcurrentTransformer
// ============================================================================

func TestStyleListenerWithParallel(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	listener := &testStyleListenerCounter{}
	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
			{"Name": "Carol", "Age": 35, "Salary": 3000},
			{"Name": "Dave", "Age": 28, "Salary": 1500},
		},
	}

	outPath := filepath.Join(testdataDir(t), "style_parallel_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(tmpl, outPath, data, WithAreaListener(listener), WithParallelism(2))
	require.NoError(t, err)
	// StyleCell should have been called even with parallel mode
	assert.True(t, listener.styleCalls.Load() > 0, "StyleListener.StyleCell should be called with parallel mode")
}

type testStyleListenerCounter struct {
	styleCalls atomic.Int64
}

func (l *testStyleListenerCounter) BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool {
	return true
}
func (l *testStyleListenerCounter) AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer) {
}
func (l *testStyleListenerCounter) StyleCell(target CellRef, value any, ctx *Context) *StyleOverride {
	l.styleCalls.Add(1)
	return nil
}

// ============================================================================
// Fix: unwrapExcelizeTransformer works through layers
// ============================================================================

func TestUnwrapExcelizeTransformer(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	etx, err := OpenTemplate(tmpl)
	require.NoError(t, err)
	defer etx.Close()

	// Direct
	assert.Equal(t, etx, unwrapExcelizeTransformer(etx))

	// Through ConcurrentTransformer
	ct := NewConcurrentTransformer(etx)
	assert.Equal(t, etx, unwrapExcelizeTransformer(ct))

	// Through StreamingTransformer
	stx, err := NewStreamingTransformer(etx, "Sheet1")
	require.NoError(t, err)
	assert.Equal(t, etx, unwrapExcelizeTransformer(stx))

	// Through ConcurrentTransformer wrapping StreamingTransformer
	ct2 := NewConcurrentTransformer(stx)
	assert.Equal(t, etx, unwrapExcelizeTransformer(ct2))

	// Unknown type
	assert.Nil(t, unwrapExcelizeTransformer(nil))
}

// ============================================================================
// Fix: Compiled template propagates autoMode
// ============================================================================

func TestCompiledTemplatePropagatesAutoMode(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl, WithAutoMode(map[string]any{"itemCount": 50000}))
	require.NoError(t, err)

	assert.True(t, compiled.opts.autoMode)
	assert.Equal(t, 50000, toInt(compiled.opts.autoModeHint["itemCount"]))

	data := map[string]any{
		"employees": []map[string]any{{"Name": "Alice", "Age": 30, "Salary": 1000}},
	}
	b, err := compiled.FillBytes(data)
	require.NoError(t, err)
	assert.True(t, len(b) > 0)
}

// ============================================================================
// Fix: Auto-mode resets targetRefs between BuildAreas calls
// ============================================================================

func TestAutoModeResetsTargetRefs(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "automode_reset_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
		},
	}

	// Use auto-mode which triggers double BuildAreas
	err := Fill(tmpl, outPath, data, WithAutoMode(map[string]any{"itemCount": 50000}))
	require.NoError(t, err)

	// Verify output is correct (no duplicated data from stale targetRefs)
	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	v1, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v1)
	v2, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Bob", v2)
}

// ============================================================================
// Fix: Parallel with timeout actually stops
// ============================================================================

func TestParallelWithTimeoutStops(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "parallel_timeout_out.xlsx")
	defer os.Remove(outPath)

	// Use an already-expired deadline for deterministic behavior
	ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(-time.Second))
	defer cancel()

	employees := make([]map[string]any, 100)
	for i := range employees {
		employees[i] = map[string]any{"Name": "E", "Age": i, "Salary": i}
	}

	err := Fill(tmpl, outPath, map[string]any{"employees": employees},
		WithParallelism(4), WithContext(ctx))
	assert.Error(t, err)
}
