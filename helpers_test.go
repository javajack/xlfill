package xlfill_test

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/javajack/xlfill"
	"github.com/xuri/excelize/v2"
)

// --- StructSliceToData ---

func TestStructSliceToData_Basic(t *testing.T) {
	type Person struct {
		Name string
		Age  int
	}
	items := []Person{
		{Name: "Alice", Age: 30},
		{Name: "Bob", Age: 25},
	}
	result := xlfill.StructSliceToData(items)
	if len(result) != 2 {
		t.Fatalf("expected 2 items, got %d", len(result))
	}
	if result[0]["Name"] != "Alice" {
		t.Errorf("expected Name=Alice, got %v", result[0]["Name"])
	}
	if result[0]["Age"] != 30 {
		t.Errorf("expected Age=30, got %v", result[0]["Age"])
	}
	if result[1]["Name"] != "Bob" {
		t.Errorf("expected Name=Bob, got %v", result[1]["Name"])
	}
}

func TestStructSliceToData_Pointer(t *testing.T) {
	type Item struct {
		ID    int
		Label string
	}
	items := []*Item{
		{ID: 1, Label: "one"},
		{ID: 2, Label: "two"},
	}
	result := xlfill.StructSliceToData(items)
	if len(result) != 2 {
		t.Fatalf("expected 2 items, got %d", len(result))
	}
	if result[0]["ID"] != 1 {
		t.Errorf("expected ID=1, got %v", result[0]["ID"])
	}
	if result[1]["Label"] != "two" {
		t.Errorf("expected Label=two, got %v", result[1]["Label"])
	}
}

func TestStructSliceToData_Nested(t *testing.T) {
	type Address struct {
		City string
	}
	type User struct {
		Name    string
		Address Address
	}
	items := []User{
		{Name: "Eve", Address: Address{City: "NYC"}},
	}
	result := xlfill.StructSliceToData(items)
	if len(result) != 1 {
		t.Fatalf("expected 1 item, got %d", len(result))
	}
	addr, ok := result[0]["Address"].(Address)
	if !ok {
		t.Fatalf("expected Address to be Address struct, got %T", result[0]["Address"])
	}
	if addr.City != "NYC" {
		t.Errorf("expected City=NYC, got %v", addr.City)
	}
}

func TestStructSliceToData_UnexportedFields(t *testing.T) {
	type Thing struct {
		Public  string
		private string //nolint:unused
	}
	items := []Thing{{Public: "yes"}}
	result := xlfill.StructSliceToData(items)
	if _, ok := result[0]["private"]; ok {
		t.Error("unexported field should not be in map")
	}
	if result[0]["Public"] != "yes" {
		t.Errorf("expected Public=yes, got %v", result[0]["Public"])
	}
}

func TestStructSliceToData_NonStruct(t *testing.T) {
	items := []int{1, 2, 3}
	result := xlfill.StructSliceToData(items)
	if len(result) != 3 {
		t.Fatalf("expected 3, got %d", len(result))
	}
	if result[0]["value"] != 1 {
		t.Errorf("expected value=1, got %v", result[0]["value"])
	}
}

func TestStructSliceToData_Empty(t *testing.T) {
	type X struct{ A int }
	result := xlfill.StructSliceToData([]X{})
	if len(result) != 0 {
		t.Errorf("expected empty, got %d", len(result))
	}
}

// --- JSONToData ---

func TestJSONToData_Valid(t *testing.T) {
	input := `{"name":"Alice","items":[1,2,3],"nested":{"key":"val"}}`
	data, err := xlfill.JSONToData(strings.NewReader(input))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if data["name"] != "Alice" {
		t.Errorf("expected name=Alice, got %v", data["name"])
	}
	items, ok := data["items"].([]any)
	if !ok || len(items) != 3 {
		t.Errorf("expected items array of length 3, got %v", data["items"])
	}
	nested, ok := data["nested"].(map[string]any)
	if !ok || nested["key"] != "val" {
		t.Errorf("expected nested.key=val, got %v", data["nested"])
	}
}

func TestJSONToData_Invalid(t *testing.T) {
	_, err := xlfill.JSONToData(strings.NewReader("not json"))
	if err == nil {
		t.Fatal("expected error for invalid JSON")
	}
	if !strings.Contains(err.Error(), "json decode") {
		t.Errorf("expected json decode error, got: %v", err)
	}
}

func TestJSONToData_Array(t *testing.T) {
	_, err := xlfill.JSONToData(strings.NewReader(`[1,2,3]`))
	if err == nil {
		t.Fatal("expected error for JSON array (not object)")
	}
}

func TestJSONToData_Empty(t *testing.T) {
	_, err := xlfill.JSONToData(strings.NewReader(""))
	if err == nil {
		t.Fatal("expected error for empty input")
	}
}

// --- SQLRowsToData ---

type mockRows struct {
	cols    []string
	data    [][]any
	cursor  int
	scanErr error
}

func (m *mockRows) Columns() ([]string, error) { return m.cols, nil }
func (m *mockRows) Next() bool {
	if m.cursor >= len(m.data) {
		return false
	}
	m.cursor++
	return true
}
func (m *mockRows) Scan(dest ...any) error {
	if m.scanErr != nil {
		return m.scanErr
	}
	row := m.data[m.cursor-1]
	for i, v := range row {
		*(dest[i].(*any)) = v
	}
	return nil
}

func TestSQLRowsToData(t *testing.T) {
	rows := &mockRows{
		cols: []string{"id", "name", "score"},
		data: [][]any{
			{1, "Alice", 95.5},
			{2, "Bob", 88.0},
		},
	}
	result, err := xlfill.SQLRowsToData(rows)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(result))
	}
	if result[0]["id"] != 1 {
		t.Errorf("expected id=1, got %v", result[0]["id"])
	}
	if result[0]["name"] != "Alice" {
		t.Errorf("expected name=Alice, got %v", result[0]["name"])
	}
	if result[1]["score"] != 88.0 {
		t.Errorf("expected score=88.0, got %v", result[1]["score"])
	}
}

func TestSQLRowsToData_Empty(t *testing.T) {
	rows := &mockRows{cols: []string{"a"}, data: nil}
	result, err := xlfill.SQLRowsToData(rows)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result) != 0 {
		t.Errorf("expected 0 rows, got %d", len(result))
	}
}

func TestSQLRowsToData_ScanError(t *testing.T) {
	rows := &mockRows{
		cols:    []string{"a"},
		data:    [][]any{{1}},
		scanErr: fmt.Errorf("scan failure"),
	}
	_, err := xlfill.SQLRowsToData(rows)
	if err == nil {
		t.Fatal("expected error")
	}
	if !strings.Contains(err.Error(), "scan row") {
		t.Errorf("expected scan row error, got: %v", err)
	}
}

// --- HTTPHandler ---

func createTestTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${title}")
	f.SetCellValue(sheet, "B1", "${value}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "test",
		Text: `jx:area(lastCell="B1")`,
	})
	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "tmpl.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	f.Close()
	return path
}

func TestHTTPHandler(t *testing.T) {
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath)
	if err != nil {
		t.Fatalf("compile: %v", err)
	}

	handler := xlfill.HTTPHandler(compiled, func(r *http.Request) (map[string]any, string, error) {
		return map[string]any{"title": "Report", "value": 42}, "myreport", nil
	})

	req := httptest.NewRequest("GET", "/download", nil)
	rec := httptest.NewRecorder()
	handler(rec, req)

	resp := rec.Result()
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		t.Errorf("expected 200, got %d", resp.StatusCode)
	}
	ct := resp.Header.Get("Content-Type")
	if ct != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" {
		t.Errorf("unexpected Content-Type: %s", ct)
	}
	cd := resp.Header.Get("Content-Disposition")
	if !strings.Contains(cd, "myreport.xlsx") {
		t.Errorf("unexpected Content-Disposition: %s", cd)
	}

	// Verify the output is a valid xlsx
	body, _ := io.ReadAll(resp.Body)
	if len(body) == 0 {
		t.Fatal("empty response body")
	}
	outFile, err := excelize.OpenReader(bytes.NewReader(body))
	if err != nil {
		t.Fatalf("response is not valid xlsx: %v", err)
	}
	defer outFile.Close()
	val, _ := outFile.GetCellValue("Sheet1", "A1")
	if val != "Report" {
		t.Errorf("expected A1=Report, got %q", val)
	}
}

func TestHTTPHandler_DefaultFilename(t *testing.T) {
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath)
	if err != nil {
		t.Fatalf("compile: %v", err)
	}

	handler := xlfill.HTTPHandler(compiled, func(r *http.Request) (map[string]any, string, error) {
		return map[string]any{"title": "X", "value": 1}, "", nil
	})

	req := httptest.NewRequest("GET", "/download", nil)
	rec := httptest.NewRecorder()
	handler(rec, req)
	resp := rec.Result()
	defer resp.Body.Close()

	cd := resp.Header.Get("Content-Disposition")
	if !strings.Contains(cd, "report.xlsx") {
		t.Errorf("expected default filename report.xlsx, got: %s", cd)
	}
}

func TestHTTPHandler_DataFnError(t *testing.T) {
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath)
	if err != nil {
		t.Fatalf("compile: %v", err)
	}

	handler := xlfill.HTTPHandler(compiled, func(r *http.Request) (map[string]any, string, error) {
		return nil, "", fmt.Errorf("db connection failed")
	})

	req := httptest.NewRequest("GET", "/download", nil)
	rec := httptest.NewRecorder()
	handler(rec, req)
	resp := rec.Result()
	defer resp.Body.Close()

	if resp.StatusCode != 500 {
		t.Errorf("expected 500, got %d", resp.StatusCode)
	}
	body, _ := io.ReadAll(resp.Body)
	if !strings.Contains(string(body), "db connection failed") {
		t.Errorf("expected error in body, got: %s", body)
	}
}

// --- FillBatch ---

func TestFillBatch(t *testing.T) {
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath)
	if err != nil {
		t.Fatalf("compile: %v", err)
	}

	tmpDir := t.TempDir()
	items := []map[string]any{
		{"title": "Report A", "value": 10},
		{"title": "Report B", "value": 20},
		{"title": "Report C", "value": 30},
	}

	err = compiled.FillBatch(items, func(i int, data map[string]any) string {
		return filepath.Join(tmpDir, fmt.Sprintf("output_%d.xlsx", i))
	})
	if err != nil {
		t.Fatalf("FillBatch: %v", err)
	}

	// Verify all 3 files
	for i, item := range items {
		path := filepath.Join(tmpDir, fmt.Sprintf("output_%d.xlsx", i))
		f, err := excelize.OpenFile(path)
		if err != nil {
			t.Fatalf("open output %d: %v", i, err)
		}
		val, _ := f.GetCellValue("Sheet1", "A1")
		expected := item["title"].(string)
		if val != expected {
			t.Errorf("output %d: expected A1=%q, got %q", i, expected, val)
		}
		f.Close()
	}
}

func TestFillBatch_Empty(t *testing.T) {
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath)
	if err != nil {
		t.Fatalf("compile: %v", err)
	}
	err = compiled.FillBatch(nil, func(i int, data map[string]any) string { return "" })
	if err != nil {
		t.Fatalf("expected nil error for empty batch, got: %v", err)
	}
}

// --- WithDocumentProperties ---

func TestWithDocumentProperties(t *testing.T) {
	tmplPath := createTestTemplate(t)
	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "docprops.xlsx")

	err := xlfill.Fill(tmplPath, outPath, map[string]any{
		"title": "Hello",
		"value": 99,
	}, xlfill.WithDocumentProperties(map[string]string{
		"title":       "My Report",
		"creator":     "xlfill",
		"subject":     "Test Subject",
		"description": "A test doc",
		"keywords":    "go,excel",
		"category":    "Reports",
	}))
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	f, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer f.Close()

	props, err := f.GetDocProps()
	if err != nil {
		t.Fatalf("GetDocProps: %v", err)
	}
	if props.Title != "My Report" {
		t.Errorf("expected title=My Report, got %q", props.Title)
	}
	if props.Creator != "xlfill" {
		t.Errorf("expected creator=xlfill, got %q", props.Creator)
	}
	if props.Subject != "Test Subject" {
		t.Errorf("expected subject=Test Subject, got %q", props.Subject)
	}
	if props.Description != "A test doc" {
		t.Errorf("expected description=A test doc, got %q", props.Description)
	}
	if props.Keywords != "go,excel" {
		t.Errorf("expected keywords=go,excel, got %q", props.Keywords)
	}
	if props.Category != "Reports" {
		t.Errorf("expected category=Reports, got %q", props.Category)
	}
}

func TestWithDocumentProperties_ViaCompiled(t *testing.T) {
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath, xlfill.WithDocumentProperties(map[string]string{
		"title":   "Compiled Report",
		"creator": "compiled-test",
	}))
	if err != nil {
		t.Fatalf("compile: %v", err)
	}

	var buf bytes.Buffer
	err = compiled.FillWriter(map[string]any{"title": "Hi", "value": 1}, &buf)
	if err != nil {
		t.Fatalf("FillWriter: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer f.Close()

	props, err := f.GetDocProps()
	if err != nil {
		t.Fatalf("GetDocProps: %v", err)
	}
	if props.Title != "Compiled Report" {
		t.Errorf("expected title=Compiled Report, got %q", props.Title)
	}
	if props.Creator != "compiled-test" {
		t.Errorf("expected creator=compiled-test, got %q", props.Creator)
	}
}

// --- Data Validation Preservation ---

func TestDataValidationPreservation(t *testing.T) {
	// Create a template with a data validation (dropdown) on a cell in the each row
	f := excelize.NewFile()
	sheet := "Sheet1"

	// Header row
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Status")

	// Data row with expressions
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Status}")

	// Add data validation (dropdown) on B2
	dv := excelize.NewDataValidation(true)
	dv.Sqref = "B2"
	if err := dv.SetDropList([]string{"Active", "Inactive", "Pending"}); err != nil {
		t.Fatal(err)
	}
	if err := f.AddDataValidation(sheet, dv); err != nil {
		t.Fatal(err)
	}

	// JXLS-style commands
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "test",
		Text: `jx:area(lastCell="B2")`,
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "test",
		Text: `jx:each(items="employees" var="e" lastCell="B2")`,
	})

	tmpDir := t.TempDir()
	tmplPath := filepath.Join(tmpDir, "dv_template.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		t.Fatal(err)
	}
	f.Close()

	// Fill with 3 items
	outPath := filepath.Join(tmpDir, "dv_output.xlsx")
	err := xlfill.Fill(tmplPath, outPath, map[string]any{
		"employees": []any{
			map[string]any{"Name": "Alice", "Status": "Active"},
			map[string]any{"Name": "Bob", "Status": "Inactive"},
			map[string]any{"Name": "Carol", "Status": "Pending"},
		},
	})
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	// Open output and check data validations
	out, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()

	dvs, err := out.GetDataValidations(sheet)
	if err != nil {
		t.Fatalf("GetDataValidations: %v", err)
	}

	// We expect at least 3 validation rules (one per row B2, B3, B4)
	// The original B2 is remapped, and we get additional B3, B4
	if len(dvs) < 3 {
		t.Errorf("expected at least 3 data validations, got %d", len(dvs))
		for i, d := range dvs {
			t.Logf("  dv[%d]: sqref=%s type=%s", i, d.Sqref, d.Type)
		}
	}

	// Check that each output row (B2, B3, B4) has a validation
	expectedCells := map[string]bool{"B2": false, "B3": false, "B4": false}
	for _, d := range dvs {
		// Check if this validation covers any of our expected cells
		for cell := range expectedCells {
			if strings.Contains(d.Sqref, cell) {
				expectedCells[cell] = true
			}
		}
	}
	for cell, found := range expectedCells {
		if !found {
			t.Errorf("expected data validation covering cell %s, not found", cell)
		}
	}
}

// --- WithStreamingSheets ---

func TestWithStreamingSheets(t *testing.T) {
	// Create template
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${title}")
	f.SetCellValue(sheet, "B1", "${value}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "test",
		Text: `jx:area(lastCell="B1")`,
	})

	tmpDir := t.TempDir()
	tmplPath := filepath.Join(tmpDir, "streaming_tmpl.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		t.Fatal(err)
	}
	f.Close()

	outPath := filepath.Join(tmpDir, "streaming_out.xlsx")
	err := xlfill.Fill(tmplPath, outPath, map[string]any{
		"title": "Streamed",
		"value": 100,
	}, xlfill.WithStreamingSheets("Sheet1"))
	if err != nil {
		t.Fatalf("Fill with streaming sheets: %v", err)
	}

	out, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()
	val, _ := out.GetCellValue(sheet, "A1")
	if val != "Streamed" {
		t.Errorf("expected A1=Streamed, got %q", val)
	}
}

func TestWithStreamingSheets_NonMatchingSheet(t *testing.T) {
	// When the named streaming sheet doesn't exist, the sheet is skipped with
	// a warning and the existing sheet is processed through the non-streaming
	// path. Output must still be correct.
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${title}")
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "test",
		Text: `jx:area(lastCell="A1")`,
	})

	tmpDir := t.TempDir()
	tmplPath := filepath.Join(tmpDir, "no_match_tmpl.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		t.Fatal(err)
	}
	f.Close()

	outPath := filepath.Join(tmpDir, "no_match_out.xlsx")
	err := xlfill.Fill(tmplPath, outPath, map[string]any{
		"title": "Normal",
	}, xlfill.WithStreamingSheets("NonExistentSheet"))
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	out, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()
	val, _ := out.GetCellValue(sheet, "A1")
	if val != "Normal" {
		t.Errorf("expected A1=Normal, got %q", val)
	}
}

// TestWithStreamingSheets_SecondSheet verifies that a sheet other than Sheet1
// can be streamed — previously a single-sheet bug made only sheets[0]
// eligible for streaming.
func TestWithStreamingSheets_SecondSheet(t *testing.T) {
	f := excelize.NewFile()
	// Sheet1 stays as the default sheet; we add a second sheet and stream it.
	f.SetCellValue("Sheet1", "A1", "static")
	_, err := f.NewSheet("Report")
	if err != nil {
		t.Fatal(err)
	}
	f.SetCellValue("Report", "A1", "Header")
	f.SetCellValue("Report", "A2", "${e.name}")
	f.AddComment("Report", excelize.Comment{
		Cell: "A1", Author: "test", Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment("Report", excelize.Comment{
		Cell: "A2", Author: "test",
		Text: `jx:each(items="employees" var="e" lastCell="A2")`,
	})

	tmpDir := t.TempDir()
	tmplPath := filepath.Join(tmpDir, "second_sheet_tmpl.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		t.Fatal(err)
	}
	f.Close()

	outPath := filepath.Join(tmpDir, "second_sheet_out.xlsx")
	err = xlfill.Fill(tmplPath, outPath, map[string]any{
		"employees": []any{
			map[string]any{"name": "Alice"},
			map[string]any{"name": "Bob"},
		},
	}, xlfill.WithStreamingSheets("Report"))
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	out, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()

	v2, _ := out.GetCellValue("Report", "A2")
	if v2 != "Alice" {
		t.Errorf("expected Report!A2=Alice, got %q", v2)
	}
	v3, _ := out.GetCellValue("Report", "A3")
	if v3 != "Bob" {
		t.Errorf("expected Report!A3=Bob, got %q", v3)
	}
}

// TestWithStreamingSheets_HyperlinkDroppedAsWarning verifies that a hyperlink
// in a streamed sheet is dropped (display text written, link omitted) and
// the drop is surfaced as a non-fatal warning that callers can inspect.
func TestWithStreamingSheets_HyperlinkDroppedAsWarning(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", `${hyperlink(e.url, e.name)}`)
	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "test", Text: `jx:area(lastCell="A1")`,
	})

	tmpDir := t.TempDir()
	tmplPath := filepath.Join(tmpDir, "hyperlink_stream.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		t.Fatal(err)
	}
	f.Close()

	filler := xlfill.NewFiller(
		xlfill.WithTemplate(tmplPath),
		xlfill.WithStreaming(true),
	)
	data := map[string]any{
		"e": map[string]any{"url": "https://example.com", "name": "Click me"},
	}
	var buf bytes.Buffer
	if err := filler.FillWriter(data, &buf); err != nil {
		t.Fatalf("FillWriter: %v", err)
	}

	warnings := filler.Warnings()
	if len(warnings) == 0 {
		t.Fatalf("expected hyperlink drop warning, got none")
	}
	found := false
	for _, w := range warnings {
		if strings.Contains(w.Message, "hyperlink dropped") {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("warnings did not include hyperlink drop notice: %v", warnings)
	}

	out, err := excelize.OpenReader(&buf)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()

	// Display text should land in the cell even though the link itself was dropped.
	v, _ := out.GetCellValue(sheet, "A1")
	if v != "Click me" {
		t.Errorf("expected display text in cell, got %q", v)
	}
}

// TestWithStreamingSheets_MultipleSheets exercises streaming two sheets at
// once and confirms both stream correctly.
func TestWithStreamingSheets_MultipleSheets(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Header A")
	f.SetCellValue("Sheet1", "A2", "${a.name}")
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A1", Author: "test", Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment("Sheet1", excelize.Comment{
		Cell: "A2", Author: "test",
		Text: `jx:each(items="listA" var="a" lastCell="A2")`,
	})
	_, err := f.NewSheet("Sheet2")
	if err != nil {
		t.Fatal(err)
	}
	f.SetCellValue("Sheet2", "A1", "Header B")
	f.SetCellValue("Sheet2", "A2", "${b.value}")
	f.AddComment("Sheet2", excelize.Comment{
		Cell: "A1", Author: "test", Text: `jx:area(lastCell="A2")`,
	})
	f.AddComment("Sheet2", excelize.Comment{
		Cell: "A2", Author: "test",
		Text: `jx:each(items="listB" var="b" lastCell="A2")`,
	})

	tmpDir := t.TempDir()
	tmplPath := filepath.Join(tmpDir, "multi_tmpl.xlsx")
	if err := f.SaveAs(tmplPath); err != nil {
		t.Fatal(err)
	}
	f.Close()

	outPath := filepath.Join(tmpDir, "multi_out.xlsx")
	err = xlfill.Fill(tmplPath, outPath, map[string]any{
		"listA": []any{
			map[string]any{"name": "X"},
			map[string]any{"name": "Y"},
		},
		"listB": []any{
			map[string]any{"value": 1},
			map[string]any{"value": 2},
		},
	}, xlfill.WithStreamingSheets("Sheet1", "Sheet2"))
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	out, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open output: %v", err)
	}
	defer out.Close()

	a2, _ := out.GetCellValue("Sheet1", "A2")
	if a2 != "X" {
		t.Errorf("expected Sheet1!A2=X, got %q", a2)
	}
	a3, _ := out.GetCellValue("Sheet1", "A3")
	if a3 != "Y" {
		t.Errorf("expected Sheet1!A3=Y, got %q", a3)
	}
	b2, _ := out.GetCellValue("Sheet2", "A2")
	if b2 != "1" {
		t.Errorf("expected Sheet2!A2=1, got %q", b2)
	}
	b3, _ := out.GetCellValue("Sheet2", "A3")
	if b3 != "2" {
		t.Errorf("expected Sheet2!A3=2, got %q", b3)
	}
}

// --- Additional edge case tests ---

func TestJSONToData_Roundtrip(t *testing.T) {
	original := map[string]any{
		"title": "Test",
		"count": float64(42), // JSON numbers are float64
	}
	b, _ := json.Marshal(original)
	data, err := xlfill.JSONToData(bytes.NewReader(b))
	if err != nil {
		t.Fatalf("JSONToData: %v", err)
	}
	if data["title"] != "Test" {
		t.Errorf("expected title=Test, got %v", data["title"])
	}
	if data["count"] != float64(42) {
		t.Errorf("expected count=42, got %v", data["count"])
	}
}

func TestFillBatch_ErrorPropagation(t *testing.T) {
	// Test that an error in one item stops the batch and propagates
	tmplPath := createTestTemplate(t)
	compiled, err := xlfill.Compile(tmplPath)
	if err != nil {
		t.Fatalf("compile: %v", err)
	}

	items := []map[string]any{
		{"title": "OK", "value": 1},
		{"title": "OK2", "value": 2},
	}

	// Use an invalid output path for the second item
	err = compiled.FillBatch(items, func(i int, data map[string]any) string {
		if i == 1 {
			return "/nonexistent/dir/should/fail.xlsx"
		}
		return filepath.Join(os.TempDir(), fmt.Sprintf("batch_test_%d.xlsx", i))
	})
	if err == nil {
		t.Fatal("expected error for invalid output path")
	}
	if !strings.Contains(err.Error(), "batch item 1") {
		t.Errorf("expected batch item 1 error, got: %v", err)
	}

	// Clean up the first file that was successfully created
	os.Remove(filepath.Join(os.TempDir(), "batch_test_0.xlsx"))
}

func TestWithDocumentProperties_EmptyMap(t *testing.T) {
	tmplPath := createTestTemplate(t)
	tmpDir := t.TempDir()
	outPath := filepath.Join(tmpDir, "empty_props.xlsx")

	err := xlfill.Fill(tmplPath, outPath, map[string]any{
		"title": "X",
		"value": 1,
	}, xlfill.WithDocumentProperties(map[string]string{}))
	if err != nil {
		t.Fatalf("Fill: %v", err)
	}

	f, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer f.Close()
	_, err = f.GetDocProps()
	if err != nil {
		t.Fatalf("GetDocProps should not error: %v", err)
	}
}

type mockColumnsError struct{}

func (m *mockColumnsError) Columns() ([]string, error) { return nil, fmt.Errorf("columns error") }
func (m *mockColumnsError) Next() bool                 { return false }
func (m *mockColumnsError) Scan(dest ...any) error     { return nil }

func TestSQLRowsToData_ColumnsError(t *testing.T) {
	_, err := xlfill.SQLRowsToData(&mockColumnsError{})
	if err == nil {
		t.Fatal("expected error")
	}
	if !strings.Contains(err.Error(), "get columns") {
		t.Errorf("expected get columns error, got: %v", err)
	}
}
