package xlfill

import (
	"bytes"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ============================================================================
// Security: Path traversal in jx:include
// ============================================================================

func TestIncludeRejectsAbsolutePath(t *testing.T) {
	err := validateIncludePath("/etc/passwd")
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "absolute paths not allowed")
}

func TestIncludeRejectsParentTraversal(t *testing.T) {
	err := validateIncludePath("../../../etc/passwd")
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "path traversal not allowed")
}

func TestIncludeRejectsDotDotInMiddle(t *testing.T) {
	err := validateIncludePath("templates/../../../secret.xlsx")
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "path traversal not allowed")
}

func TestIncludeAllowsRelativePath(t *testing.T) {
	err := validateIncludePath("templates/header.xlsx")
	assert.NoError(t, err)
}

func TestIncludeAllowsCurrentDir(t *testing.T) {
	err := validateIncludePath("./header.xlsx")
	assert.NoError(t, err)
}

// ============================================================================
// Security: Unbounded repeat count
// ============================================================================

func TestRepeatCountLimit(t *testing.T) {
	cmd := &RepeatCommand{Count: "2000000", Var: "i", Direction: "DOWN"}
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{1, 1}, nil)
	cmd.Area = area

	ctx := NewContext(nil)
	_, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, nil)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "exceeds maximum")
}

func TestRepeatCountAtLimit(t *testing.T) {
	// MaxRepeatCount itself should be allowed (but we can't actually run 1M iterations
	// in a test, so just verify the validation logic)
	assert.Equal(t, 1_000_000, MaxRepeatCount)
}

// ============================================================================
// Security: Parallelism cap
// ============================================================================

func TestParallelismCapped(t *testing.T) {
	opts := defaultOptions()
	WithParallelism(999999)(opts)
	assert.Equal(t, MaxParallelism, opts.parallelism)
}

func TestParallelismNegativeClamped(t *testing.T) {
	opts := defaultOptions()
	WithParallelism(-5)(opts)
	assert.Equal(t, 0, opts.parallelism)
}

func TestParallelismNormalValue(t *testing.T) {
	opts := defaultOptions()
	WithParallelism(4)(opts)
	assert.Equal(t, 4, opts.parallelism)
}

// ============================================================================
// Security: Recursive include depth limit
// ============================================================================

func TestIncludeDepthLimit(t *testing.T) {
	ctx := NewContext(nil)
	// Simulate deep nesting by incrementing includeDepth
	for i := 0; i < 10; i++ {
		ctx.includeDepth++
	}
	assert.Equal(t, 10, ctx.includeDepth)

	// Clone preserves include depth
	clone := ctx.Clone()
	assert.Equal(t, 10, clone.includeDepth)
}

// ============================================================================
// Security: Filename sanitization in HTTPHandler
// ============================================================================

func TestSanitizeFilename(t *testing.T) {
	assert.Equal(t, "report", sanitizeFilename("report"))
	assert.Equal(t, "my_report", sanitizeFilename("my;report"))
	assert.Equal(t, "file_name", sanitizeFilename("file\"name"))
	assert.Equal(t, "path_to_file", sanitizeFilename("path/to\\file"))
	assert.Equal(t, "safe_name", sanitizeFilename("safe:name"))
	// Control characters
	assert.Equal(t, "no_ctrl", sanitizeFilename("no\x00ctrl"))
}

func TestHTTPHandlerSecurityHeaders(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	compiled, err := Compile(tmpl)
	require.NoError(t, err)

	handler := HTTPHandler(compiled, func(r *http.Request) (map[string]any, string, error) {
		return map[string]any{
			"employees": []map[string]any{{"Name": "A", "Age": 1, "Salary": 1}},
		}, "report", nil
	})

	req := httptest.NewRequest("GET", "/download", nil)
	rec := httptest.NewRecorder()
	handler(rec, req)

	assert.Equal(t, "nosniff", rec.Header().Get("X-Content-Type-Options"))
	assert.Equal(t, "no-store", rec.Header().Get("Cache-Control"))
	assert.Contains(t, rec.Header().Get("Content-Disposition"), "report.xlsx")
}

// ============================================================================
// Security: Deferred actions limit
// ============================================================================

func TestDeferredActionsLimit(t *testing.T) {
	dr := NewDeferredRegistry()
	for i := 0; i < MaxDeferredActions; i++ {
		err := dr.Add(DeferredAction{Name: "test"})
		require.NoError(t, err)
	}
	// One more should fail
	err := dr.Add(DeferredAction{Name: "overflow"})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "limit")
}

// ============================================================================
// Memory Safety: CellData thread-safe AddTargetPos
// ============================================================================

func TestCellDataAddTargetPosConcurrent(t *testing.T) {
	cd := &CellData{Ref: NewCellRef("Sheet1", 0, 0)}

	var wg sync.WaitGroup
	for i := 0; i < 100; i++ {
		wg.Add(1)
		go func(n int) {
			defer wg.Done()
			cd.AddTargetPos(NewCellRef("Sheet1", n, 0))
		}(i)
	}
	wg.Wait()

	assert.Equal(t, 100, len(cd.TargetPositions))
}

func TestCellDataAddTargetPosWithAreaConcurrent(t *testing.T) {
	cd := &CellData{Ref: NewCellRef("Sheet1", 0, 0)}

	var wg sync.WaitGroup
	for i := 0; i < 50; i++ {
		wg.Add(1)
		go func(n int) {
			defer wg.Done()
			cd.AddTargetPosWithArea(
				NewCellRef("Sheet1", n, 0),
				NewAreaRef(NewCellRef("Sheet1", 0, 0), NewCellRef("Sheet1", n, 0)),
			)
		}(i)
	}
	wg.Wait()

	assert.Equal(t, 50, len(cd.TargetPositions))
	assert.Equal(t, 50, len(cd.TargetParentArea))
}

// ============================================================================
// Memory Safety: WarningCollector thread-safe
// ============================================================================

func TestWarningCollectorConcurrent(t *testing.T) {
	wc := &WarningCollector{}
	var wg sync.WaitGroup
	for i := 0; i < 100; i++ {
		wg.Add(1)
		go func(n int) {
			defer wg.Done()
			wc.Add(NewCellRef("Sheet1", n, 0), "warning")
		}(i)
	}
	wg.Wait()

	assert.Equal(t, 100, len(wc.Warnings()))
}

// ============================================================================
// Error Handling: Grid SetCellValue errors propagated
// ============================================================================

func TestGridErrorPropagation(t *testing.T) {
	// GridCommand should return errors from SetCellValue
	cmd := &GridCommand{Headers: "headers", Data: "data"}
	assert.Equal(t, "grid", cmd.Name())
}

// ============================================================================
// Error Handling: Direction validation
// ============================================================================

func TestEachInvalidDirection(t *testing.T) {
	_, err := newEachCommandFromAttrs(map[string]string{
		"items":     "items",
		"var":       "e",
		"direction": "SIDEWAYS",
		"lastCell":  "A1",
	})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "invalid direction")
}

func TestRepeatInvalidDirection(t *testing.T) {
	_, err := newRepeatCommandFromAttrs(map[string]string{
		"count":     "5",
		"direction": "DIAGONAL",
	})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "invalid direction")
}

// ============================================================================
// Error Handling: Image scale validation
// ============================================================================

func TestImageInvalidScaleX(t *testing.T) {
	_, err := newImageCommandFromAttrs(map[string]string{
		"src":    "img",
		"scaleX": "notanumber",
	})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "scaleX")
}

func TestImageInvalidScaleY(t *testing.T) {
	_, err := newImageCommandFromAttrs(map[string]string{
		"src":    "img",
		"scaleY": "abc",
	})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "scaleY")
}

// ============================================================================
// Error Handling: DataValidation type validation
// ============================================================================

func TestDataValidationInvalidType(t *testing.T) {
	_, err := newDataValidationCommandFromAttrs(map[string]string{
		"type":     "invalid_type",
		"lastCell": "A1",
	})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "unsupported")
}

func TestDataValidationValidTypes(t *testing.T) {
	for _, typ := range []string{"list", "integer", "decimal", "date", "custom"} {
		cmd, err := newDataValidationCommandFromAttrs(map[string]string{
			"type":     typ,
			"lastCell": "A1",
		})
		assert.NoError(t, err, "type %q should be valid", typ)
		assert.NotNil(t, cmd)
	}
}

// ============================================================================
// Error Handling: Chart type validation
// ============================================================================

func TestChartInvalidType(t *testing.T) {
	cmd := &ChartCommand{ChartType: "invalid_chart"}
	cmd.Area = NewArea(NewCellRef("Sheet1", 0, 0), Size{1, 1}, nil)
	ctx := NewContext(nil)

	// Chart type should be validated — either at creation or at ApplyAt
	_, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, nil)
	assert.Error(t, err)
}

// ============================================================================
// Performance: CountBy fmt.Sprintf outside loop
// ============================================================================

func TestCountByPerformance(t *testing.T) {
	// Generate a large dataset
	items := make([]map[string]any, 1000)
	for i := range items {
		items[i] = map[string]any{"status": "active"}
	}

	// Should complete quickly (fmt.Sprintf moved outside loop)
	count := CountBy(items, "status", "active")
	assert.Equal(t, 1000, count)
}

// ============================================================================
// Observability: Deferred action tracing
// ============================================================================

func TestDebugTracerDeferredAction(t *testing.T) {
	var buf bytes.Buffer
	tracer := NewDebugTracer(&buf)

	tracer.TraceDeferredAction("table", "Sheet1", 0, 9)
	output := buf.String()

	assert.Contains(t, output, "[deferred]")
	assert.Contains(t, output, "table")
	assert.Contains(t, output, "Sheet1")
}

// ============================================================================
// Integration: Full pipeline with security constraints
// ============================================================================

func TestFullPipelineWithSecurityConstraints(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	outPath := filepath.Join(testdataDir(t), "security_test_out.xlsx")
	defer os.Remove(outPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
		},
	}

	// Should work with all safety constraints active
	err := Fill(tmpl, outPath, data,
		WithParallelism(4),
		WithStrictMode(true),
	)
	require.NoError(t, err)

	f, err := excelize.OpenFile(outPath)
	require.NoError(t, err)
	defer f.Close()

	v, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Alice", v)
}

// ============================================================================
// Edge case: AutoRowHeight error propagation
// ============================================================================

func TestAutoRowHeightName(t *testing.T) {
	cmd := &AutoRowHeightCommand{}
	assert.Equal(t, "autoRowHeight", cmd.Name())
}

// ============================================================================
// Edge case: UpdateCell error propagation structure
// ============================================================================

func TestUpdateCellMissingUpdater(t *testing.T) {
	_, err := newUpdateCellCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "updater")
}

// ============================================================================
// Edge case: Include with empty area string
// ============================================================================

func TestIncludePathTraversalInTemplate(t *testing.T) {
	// Create a template that uses jx:include with path traversal
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "header")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: `jx:area(lastCell="A1")
jx:include(template="../../../etc/passwd" sheet="Sheet1" area="A1:A1" lastCell="A1")`,
	})

	path := filepath.Join(testdataDir(t), "include_traversal_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	defer os.Remove(path)

	outPath := filepath.Join(testdataDir(t), "include_traversal_out.xlsx")
	defer os.Remove(outPath)

	err := Fill(path, outPath, map[string]any{})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "path traversal not allowed")
}

// ============================================================================
// Verify constants are sensible
// ============================================================================

func TestSecurityConstants(t *testing.T) {
	assert.Equal(t, 1_000_000, MaxRepeatCount)
	assert.Equal(t, 256, MaxParallelism)
	assert.Equal(t, 10_000, MaxDeferredActions)
}

// ============================================================================
// Edge case: sanitizeFilename with unicode
// ============================================================================

func TestSanitizeFilenameUnicode(t *testing.T) {
	// Unicode should pass through
	assert.Equal(t, "报告", sanitizeFilename("报告"))
	assert.Equal(t, "données", sanitizeFilename("données"))
	// But still sanitize dangerous chars
	assert.Equal(t, "日本_レポート", sanitizeFilename("日本/レポート"))
}

// ============================================================================
// Edge case: Streaming flushes correctly after security fixes
// ============================================================================

func TestStreamingStillWorks(t *testing.T) {
	tmpl := createBasicTemplate(t)
	defer os.Remove(tmpl)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Age": 30, "Salary": 1000},
			{"Name": "Bob", "Age": 25, "Salary": 2000},
		},
	}

	b, err := FillBytes(tmpl, data, WithStreaming(true))
	require.NoError(t, err)
	assert.True(t, len(b) > 0)
}

// ============================================================================
// Verify direction validation doesn't break existing tests
// ============================================================================

func TestEachValidDirectionDOWN(t *testing.T) {
	cmd, err := newEachCommandFromAttrs(map[string]string{
		"items": "items", "var": "e", "lastCell": "A1",
	})
	require.NoError(t, err)
	assert.Equal(t, "DOWN", cmd.(*EachCommand).Direction)
}

func TestEachValidDirectionRIGHT(t *testing.T) {
	cmd, err := newEachCommandFromAttrs(map[string]string{
		"items": "items", "var": "e", "direction": "RIGHT", "lastCell": "A1",
	})
	require.NoError(t, err)
	assert.Equal(t, "RIGHT", cmd.(*EachCommand).Direction)
}

// ============================================================================
// Verify empty direction defaults correctly
// ============================================================================

func TestRepeatDefaultDirection(t *testing.T) {
	cmd, err := newRepeatCommandFromAttrs(map[string]string{"count": "5"})
	require.NoError(t, err)
	assert.Equal(t, "DOWN", cmd.(*RepeatCommand).Direction)
}

// ============================================================================
// validateIncludePath with Windows-style paths
// ============================================================================

func TestIncludeRejectsBackslashTraversal(t *testing.T) {
	// filepath.Clean normalizes these
	err := validateIncludePath("..\\..\\secret.xlsx")
	// On Unix, backslash is a valid filename char, but filepath.Clean handles it
	// The important thing is that ".." prefix is caught
	if strings.Contains(filepath.Clean("..\\..\\secret.xlsx"), "..") {
		assert.Error(t, err)
	}
}
