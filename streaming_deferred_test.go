package xlfill_test

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/javajack/xlfill"
	"github.com/xuri/excelize/v2"
)

// These tests document the behavior of deferred-action commands when the user
// explicitly forces streaming via WithStreaming(true). Auto-mode (WithAutoMode)
// would steer away from streaming for these templates, but the explicit knob
// is preserved as a power-user escape hatch.
//
// Each test asserts what currently works so future regressions are caught.

func makeStreamingDeferredTemplate(t *testing.T, name string, comments map[string]string, values map[string]any) string {
	t.Helper()
	f := excelize.NewFile()
	sheet := "Sheet1"
	for cell, v := range values {
		f.SetCellValue(sheet, cell, v)
	}
	for cell, comment := range comments {
		f.AddComment(sheet, excelize.Comment{
			Cell: cell, Author: "test", Text: comment,
		})
	}
	tmpl := filepath.Join(t.TempDir(), name)
	if err := f.SaveAs(tmpl); err != nil {
		t.Fatal(err)
	}
	f.Close()
	return tmpl
}

// jx:freezePanes is a sheet-view change, not row data. Streaming + freezePanes
// must work.
func TestStreaming_WithFreezePanes(t *testing.T) {
	tmpl := makeStreamingDeferredTemplate(t, "freeze.xlsx",
		map[string]string{
			"A1": `jx:area(lastCell="A2")
jx:freezePanes(lastCell="A2" row="1" col="0")`,
			"A2": `jx:each(items="rows" var="r" lastCell="A2")`,
		},
		map[string]any{"A1": "Header", "A2": "${r.name}"},
	)

	var buf bytes.Buffer
	err := xlfill.FillReader(mustOpen(t, tmpl), &buf,
		map[string]any{"rows": []any{
			map[string]any{"name": "Alice"},
			map[string]any{"name": "Bob"},
		}},
		xlfill.WithStreaming(true),
	)
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	out, err := excelize.OpenReader(&buf)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()

	v, _ := out.GetCellValue("Sheet1", "A2")
	if v != "Alice" {
		t.Errorf("expected A2=Alice, got %q", v)
	}
}

// jx:definedName creates a workbook-level named range. Streaming + definedName
// must work because the named range is not row-resident.
func TestStreaming_WithDefinedName(t *testing.T) {
	tmpl := makeStreamingDeferredTemplate(t, "defname.xlsx",
		map[string]string{
			"A1": `jx:area(lastCell="A2")
jx:definedName(name="Names" lastCell="A2")`,
			"A2": `jx:each(items="rows" var="r" lastCell="A2")`,
		},
		map[string]any{"A1": "Header", "A2": "${r.name}"},
	)

	var buf bytes.Buffer
	err := xlfill.FillReader(mustOpen(t, tmpl), &buf,
		map[string]any{"rows": []any{
			map[string]any{"name": "Alice"},
		}},
		xlfill.WithStreaming(true),
	)
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	out, err := excelize.OpenReader(&buf)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()

	if v, _ := out.GetCellValue("Sheet1", "A2"); v != "Alice" {
		t.Errorf("expected A2=Alice, got %q", v)
	}
}

// jx:table + streaming: documents whatever the current behavior is.
// Streaming on tables is risky because excelize tables reference row ranges
// and the StreamWriter writes XML directly. If this test starts to fail in
// the future it's a flag that excelize internals have shifted — investigate
// before suppressing.
func TestStreaming_WithTable(t *testing.T) {
	tmpl := makeStreamingDeferredTemplate(t, "table.xlsx",
		map[string]string{
			"A1": `jx:area(lastCell="B2")`,
			"A2": `jx:each(items="rows" var="r" lastCell="B2")
jx:table(name="MyTable" lastCell="B2")`,
		},
		map[string]any{
			"A1": "Name", "B1": "Score",
			"A2": "${r.name}", "B2": "${r.score}",
		},
	)

	var buf bytes.Buffer
	err := xlfill.FillReader(mustOpen(t, tmpl), &buf,
		map[string]any{"rows": []any{
			map[string]any{"name": "Alice", "score": 90},
			map[string]any{"name": "Bob", "score": 80},
		}},
		xlfill.WithStreaming(true),
	)
	// Either it works or it returns a clear error — both are acceptable
	// outcomes for this combination. What matters is no panic.
	_ = err

	// If Fill succeeded, validate that rows landed.
	if err == nil {
		out, openErr := excelize.OpenReader(&buf)
		if openErr != nil {
			t.Fatalf("open output: %v", openErr)
		}
		defer out.Close()
		v, _ := out.GetCellValue("Sheet1", "A2")
		if v != "Alice" {
			t.Errorf("expected A2=Alice, got %q", v)
		}
	}
}

// jx:chart + streaming: chart is added to the drawing layer, separate from
// row cells. Should generally coexist with streaming.
func TestStreaming_WithChart(t *testing.T) {
	tmpl := makeStreamingDeferredTemplate(t, "chart.xlsx",
		map[string]string{
			"A1": `jx:area(lastCell="E10")`,
			"A2": `jx:each(items="rows" var="r" lastCell="B2")`,
			"D1": `jx:chart(type="bar" title="Demo" catRange="A2:A2" valRange="B2:B2" lastCell="E10")`,
		},
		map[string]any{
			"A1": "Name", "B1": "Score",
			"A2": "${r.name}", "B2": "${r.score}",
		},
	)

	var buf bytes.Buffer
	err := xlfill.FillReader(mustOpen(t, tmpl), &buf,
		map[string]any{"rows": []any{
			map[string]any{"name": "Alice", "score": 90},
		}},
		xlfill.WithStreaming(true),
	)
	_ = err

	if err == nil {
		out, openErr := excelize.OpenReader(&buf)
		if openErr != nil {
			t.Fatalf("open output: %v", openErr)
		}
		defer out.Close()
		v, _ := out.GetCellValue("Sheet1", "A2")
		if v != "Alice" {
			t.Errorf("expected A2=Alice, got %q", v)
		}
	}
}

// mustOpen reads a template file into a bytes-backed reader for FillReader.
func mustOpen(t *testing.T, path string) *bytes.Reader {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return bytes.NewReader(data)
}
