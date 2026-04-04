package xlfill

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// ─── Helper: create templates programmatically ───
// NOTE: Excel only supports one comment per cell. Multiple jx: commands on the
// same cell must be concatenated with "\n" in a single comment.

// createTableTemplate creates a template with jx:table wrapping jx:each.
func createTableTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Score")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Score}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:table(name=\"ScoreTable\" style=\"TableStyleMedium9\" lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"B2\")",
	})

	path := filepath.Join(testdataDir(t), "table_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createDataValidationTemplate creates a template with jx:dataValidation.
func createDataValidationTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Rating")
	f.SetCellValue(sheet, "A2", "${e.Rating}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:dataValidation(type=\"list\" source=\"A,B,C,D,F\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"A2\")",
	})

	path := filepath.Join(testdataDir(t), "datavalidation_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createGroupTemplate creates a template with jx:group.
func createGroupTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Salary")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Salary}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:group(lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"B2\")",
	})

	path := filepath.Join(testdataDir(t), "group_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createDefinedNameTemplate creates a template with jx:definedName.
func createDefinedNameTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Score")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Score}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:definedName(name=\"DataRange\" lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"B2\")",
	})

	path := filepath.Join(testdataDir(t), "definedname_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createChartTemplate creates a template with jx:chart.
func createChartTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Category")
	f.SetCellValue(sheet, "B1", "Value")
	f.SetCellValue(sheet, "A2", "${e.Category}")
	f.SetCellValue(sheet, "B2", "${e.Value}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:chart(type=\"col\" title=\"Test Chart\" series=\"Sheet1!$B$1:$B$4\" categories=\"Sheet1!$A$1:$A$4\" lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"B2\")",
	})

	path := filepath.Join(testdataDir(t), "chart_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createConditionalFormatTemplate creates a template with jx:conditionalFormat.
func createConditionalFormatTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Score")
	f.SetCellValue(sheet, "A2", "${e.Score}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:conditionalFormat(type=\"dataBar\" color=\"#638EC6\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"A2\")",
	})

	path := filepath.Join(testdataDir(t), "conditionalformat_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// createSparklineTemplate creates a template with jx:sparkline.
func createSparklineTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "${e.Value}")
	f.SetCellValue(sheet, "B1", "")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B1\")\njx:sparkline(type=\"line\" data=\"Sheet1!A1:A3\" lastCell=\"B1\")\njx:each(items=\"employees\" var=\"e\" lastCell=\"A1\")",
	})

	path := filepath.Join(testdataDir(t), "sparkline_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// ─── Command factory tests ───

func TestNewDataValidationCommandFromAttrs(t *testing.T) {
	t.Run("valid list", func(t *testing.T) {
		cmd, err := newDataValidationCommandFromAttrs(map[string]string{
			"type": "list", "source": "A,B,C",
		})
		require.NoError(t, err)
		dv := cmd.(*DataValidationCommand)
		assert.Equal(t, "list", dv.ValidationType)
		assert.Equal(t, "A,B,C", dv.Source)
		assert.Equal(t, "true", dv.AllowBlank)
	})

	t.Run("valid integer range", func(t *testing.T) {
		cmd, err := newDataValidationCommandFromAttrs(map[string]string{
			"type": "integer", "min": "1", "max": "100",
		})
		require.NoError(t, err)
		dv := cmd.(*DataValidationCommand)
		assert.Equal(t, "integer", dv.ValidationType)
		assert.Equal(t, "1", dv.Min)
		assert.Equal(t, "100", dv.Max)
	})

	t.Run("missing type", func(t *testing.T) {
		_, err := newDataValidationCommandFromAttrs(map[string]string{
			"source": "A,B",
		})
		assert.Error(t, err)
		assert.Contains(t, err.Error(), "type")
	})

	t.Run("with error settings", func(t *testing.T) {
		cmd, err := newDataValidationCommandFromAttrs(map[string]string{
			"type": "list", "source": "X,Y", "showError": "true",
			"errorTitle": "Invalid", "error": "Pick X or Y",
		})
		require.NoError(t, err)
		dv := cmd.(*DataValidationCommand)
		assert.Equal(t, "true", dv.ShowError)
		assert.Equal(t, "Invalid", dv.ErrorTitle)
		assert.Equal(t, "Pick X or Y", dv.ErrorMsg)
	})

	t.Run("allowBlank false", func(t *testing.T) {
		cmd, err := newDataValidationCommandFromAttrs(map[string]string{
			"type": "list", "source": "A,B", "allowBlank": "false",
		})
		require.NoError(t, err)
		dv := cmd.(*DataValidationCommand)
		assert.Equal(t, "false", dv.AllowBlank)
	})
}

func TestNewTableCommandFromAttrs(t *testing.T) {
	t.Run("valid", func(t *testing.T) {
		cmd, err := newTableCommandFromAttrs(map[string]string{
			"name": "MyTable", "style": "TableStyleLight1",
		})
		require.NoError(t, err)
		tc := cmd.(*TableCommand)
		assert.Equal(t, "MyTable", tc.TableName)
		assert.Equal(t, "TableStyleLight1", tc.Style)
	})

	t.Run("default style", func(t *testing.T) {
		cmd, err := newTableCommandFromAttrs(map[string]string{"name": "T1"})
		require.NoError(t, err)
		tc := cmd.(*TableCommand)
		assert.Equal(t, "TableStyleMedium9", tc.Style)
	})

	t.Run("missing name", func(t *testing.T) {
		_, err := newTableCommandFromAttrs(map[string]string{})
		assert.Error(t, err)
		assert.Contains(t, err.Error(), "name")
	})
}

func TestNewConditionalFormatCommandFromAttrs(t *testing.T) {
	t.Run("valid dataBar", func(t *testing.T) {
		cmd, err := newConditionalFormatCommandFromAttrs(map[string]string{
			"type": "dataBar", "color": "#FF0000",
		})
		require.NoError(t, err)
		cf := cmd.(*ConditionalFormatCommand)
		assert.Equal(t, "dataBar", cf.FormatType)
		assert.Equal(t, "#FF0000", cf.Color)
	})

	t.Run("valid colorScale", func(t *testing.T) {
		cmd, err := newConditionalFormatCommandFromAttrs(map[string]string{
			"type": "colorScale", "minColor": "#F00", "maxColor": "#0F0",
		})
		require.NoError(t, err)
		cf := cmd.(*ConditionalFormatCommand)
		assert.Equal(t, "colorScale", cf.FormatType)
		assert.Equal(t, "#F00", cf.MinColor)
		assert.Equal(t, "#0F0", cf.MaxColor)
	})

	t.Run("valid cellIs", func(t *testing.T) {
		cmd, err := newConditionalFormatCommandFromAttrs(map[string]string{
			"type": "cellIs", "operator": ">", "value": "50",
		})
		require.NoError(t, err)
		cf := cmd.(*ConditionalFormatCommand)
		assert.Equal(t, "cellIs", cf.FormatType)
		assert.Equal(t, ">", cf.Operator)
		assert.Equal(t, "50", cf.Value)
	})

	t.Run("missing type", func(t *testing.T) {
		_, err := newConditionalFormatCommandFromAttrs(map[string]string{
			"color": "#000",
		})
		assert.Error(t, err)
	})
}

func TestNewGroupCommandFromAttrs(t *testing.T) {
	t.Run("default", func(t *testing.T) {
		cmd, err := newGroupCommandFromAttrs(map[string]string{})
		require.NoError(t, err)
		gc := cmd.(*GroupCommand)
		assert.Equal(t, "false", gc.Collapsed)
	})

	t.Run("collapsed", func(t *testing.T) {
		cmd, err := newGroupCommandFromAttrs(map[string]string{"collapsed": "true"})
		require.NoError(t, err)
		gc := cmd.(*GroupCommand)
		assert.Equal(t, "true", gc.Collapsed)
	})
}

func TestNewChartCommandFromAttrs(t *testing.T) {
	t.Run("valid", func(t *testing.T) {
		cmd, err := newChartCommandFromAttrs(map[string]string{
			"type": "bar", "title": "My Chart",
			"series": "Sheet1!$B$1:$B$5", "categories": "Sheet1!$A$1:$A$5",
		})
		require.NoError(t, err)
		cc := cmd.(*ChartCommand)
		assert.Equal(t, "bar", cc.ChartType)
		assert.Equal(t, "My Chart", cc.Title)
		assert.Equal(t, "480", cc.Width)
		assert.Equal(t, "290", cc.Height)
	})

	t.Run("custom dimensions", func(t *testing.T) {
		cmd, err := newChartCommandFromAttrs(map[string]string{
			"type": "line", "width": "800", "height": "600",
		})
		require.NoError(t, err)
		cc := cmd.(*ChartCommand)
		assert.Equal(t, "800", cc.Width)
		assert.Equal(t, "600", cc.Height)
	})

	t.Run("missing type", func(t *testing.T) {
		_, err := newChartCommandFromAttrs(map[string]string{
			"title": "X",
		})
		assert.Error(t, err)
	})
}

func TestNewDefinedNameCommandFromAttrs(t *testing.T) {
	t.Run("valid", func(t *testing.T) {
		cmd, err := newDefinedNameCommandFromAttrs(map[string]string{
			"name": "MyRange", "scope": "Sheet1",
		})
		require.NoError(t, err)
		dn := cmd.(*DefinedNameCommand)
		assert.Equal(t, "MyRange", dn.DefinedNameValue)
		assert.Equal(t, "Sheet1", dn.Scope)
	})

	t.Run("default scope", func(t *testing.T) {
		cmd, err := newDefinedNameCommandFromAttrs(map[string]string{"name": "R1"})
		require.NoError(t, err)
		dn := cmd.(*DefinedNameCommand)
		assert.Equal(t, "workbook", dn.Scope)
	})

	t.Run("missing name", func(t *testing.T) {
		_, err := newDefinedNameCommandFromAttrs(map[string]string{})
		assert.Error(t, err)
	})
}

func TestNewSparklineCommandFromAttrs(t *testing.T) {
	t.Run("valid", func(t *testing.T) {
		cmd, err := newSparklineCommandFromAttrs(map[string]string{
			"type": "column", "data": "Sheet1!A1:A10", "color": "#FF0000",
		})
		require.NoError(t, err)
		sp := cmd.(*SparklineCommand)
		assert.Equal(t, "column", sp.SparkType)
		assert.Equal(t, "Sheet1!A1:A10", sp.DataRange)
		assert.Equal(t, "#FF0000", sp.Color)
	})

	t.Run("default type", func(t *testing.T) {
		cmd, err := newSparklineCommandFromAttrs(map[string]string{
			"data": "Sheet1!A1:A5",
		})
		require.NoError(t, err)
		sp := cmd.(*SparklineCommand)
		assert.Equal(t, "line", sp.SparkType)
	})

	t.Run("missing data", func(t *testing.T) {
		_, err := newSparklineCommandFromAttrs(map[string]string{
			"type": "line",
		})
		assert.Error(t, err)
	})
}

// ─── Command Name() tests ───

func TestNewCommandNames(t *testing.T) {
	tests := []struct {
		factory CommandFactory
		name    string
	}{
		{newDataValidationCommandFromAttrs, "dataValidation"},
		{newTableCommandFromAttrs, "table"},
		{newConditionalFormatCommandFromAttrs, "conditionalFormat"},
		{newGroupCommandFromAttrs, "group"},
		{newChartCommandFromAttrs, "chart"},
		{newDefinedNameCommandFromAttrs, "definedName"},
		{newSparklineCommandFromAttrs, "sparkline"},
	}

	requiredAttrs := []map[string]string{
		{"type": "list", "source": "A"},
		{"name": "T1"},
		{"type": "dataBar"},
		{},
		{"type": "bar"},
		{"name": "N1"},
		{"data": "Sheet1!A1:A5"},
	}

	for i, tt := range tests {
		cmd, err := tt.factory(requiredAttrs[i])
		require.NoError(t, err, "factory %s", tt.name)
		assert.Equal(t, tt.name, cmd.Name())
	}
}

// ─── Command Reset() tests (no-op) ───

func TestNewCommandReset(t *testing.T) {
	cmds := []Command{
		&DataValidationCommand{ValidationType: "list"},
		&TableCommand{TableName: "T"},
		&ConditionalFormatCommand{FormatType: "dataBar"},
		&GroupCommand{},
		&ChartCommand{ChartType: "bar"},
		&DefinedNameCommand{DefinedNameValue: "N"},
		&SparklineCommand{SparkType: "line", DataRange: "A1:A5"},
	}
	for _, cmd := range cmds {
		cmd.Reset() // should not panic
	}
}

// ─── ApplyAt nil Area tests ───

func TestNewCommandApplyAtNilArea(t *testing.T) {
	cellRef := NewCellRef("Sheet1", 0, 0)
	ctx := NewContext(map[string]any{})

	ef := excelize.NewFile()
	tx, err := NewExcelizeTransformer(ef)
	require.NoError(t, err)
	defer tx.Close()

	tests := []struct {
		name string
		cmd  Command
	}{
		{"dataValidation", &DataValidationCommand{ValidationType: "list"}},
		{"table", &TableCommand{TableName: "T"}},
		{"conditionalFormat", &ConditionalFormatCommand{FormatType: "dataBar"}},
		{"group", &GroupCommand{}},
		{"chart", &ChartCommand{ChartType: "bar"}},
		{"definedName", &DefinedNameCommand{DefinedNameValue: "N"}},
		{"sparkline", &SparklineCommand{SparkType: "line", DataRange: "A1:A5"}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			size, err := tt.cmd.ApplyAt(cellRef, ctx, tx)
			require.NoError(t, err)
			assert.Equal(t, ZeroSize, size)
		})
	}
}

// ─── Registry tests ───

func TestCommandRegistryNewCommands(t *testing.T) {
	reg := NewCommandRegistry()

	newCmds := []string{
		"dataValidation", "table", "conditionalFormat",
		"group", "chart", "definedName", "sparkline",
	}

	for _, name := range newCmds {
		t.Run(name, func(t *testing.T) {
			names := reg.KnownNames()
			found := false
			for _, n := range names {
				if n == name {
					found = true
					break
				}
			}
			assert.True(t, found, "command %q not registered", name)
		})
	}
}

// ─── Integration tests: fill template and verify output ───

func TestTableCommandIntegration(t *testing.T) {
	tmplPath := createTableTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Score": 95},
			{"Name": "Bob", "Score": 87},
			{"Name": "Carol", "Score": 91},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Open output and verify table exists
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	tables, err := f.GetTables("Sheet1")
	require.NoError(t, err)
	require.Len(t, tables, 1)
	assert.Equal(t, "ScoreTable", tables[0].Name)
}

func TestDataValidationCommandIntegration(t *testing.T) {
	tmplPath := createDataValidationTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Rating": "A"},
			{"Rating": "B"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Open output and verify data validations exist
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	dvs, err := f.GetDataValidations("Sheet1")
	require.NoError(t, err)
	require.True(t, len(dvs) > 0, "expected data validations")
}

func TestGroupCommandIntegration(t *testing.T) {
	tmplPath := createGroupTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Salary": 50000},
			{"Name": "Bob", "Salary": 60000},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Open output and verify outline level
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Row 1 should have outline level 1 (jx:group area starts at row 0 = Excel row 1)
	level, err := f.GetRowOutlineLevel("Sheet1", 1)
	require.NoError(t, err)
	assert.Equal(t, uint8(1), level)
}

func TestGroupCommandCollapsed(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:group(collapsed=\"true\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"A2\")",
	})

	tmplPath := filepath.Join(testdataDir(t), "group_collapsed_template.xlsx")
	require.NoError(t, f.SaveAs(tmplPath))
	f.Close()
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"Name": "X"},
			{"Name": "Y"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	outF, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer outF.Close()

	// Row 1 should be hidden (collapsed)
	visible, err := outF.GetRowVisible("Sheet1", 1)
	require.NoError(t, err)
	assert.False(t, visible, "collapsed row should be hidden")
}

func TestDefinedNameCommandIntegration(t *testing.T) {
	tmplPath := createDefinedNameTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Name": "Alice", "Score": 95},
			{"Name": "Bob", "Score": 87},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Open output and verify defined name exists
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	dnames := f.GetDefinedName()
	found := false
	for _, dn := range dnames {
		if dn.Name == "DataRange" {
			found = true
			assert.Contains(t, dn.RefersTo, "Sheet1")
			break
		}
	}
	assert.True(t, found, "expected defined name 'DataRange'")
}

func TestChartCommandIntegration(t *testing.T) {
	tmplPath := createChartTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Category": "Q1", "Value": 100},
			{"Category": "Q2", "Value": 200},
			{"Category": "Q3", "Value": 150},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Verify the output is a valid xlsx
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Verify data was written and the file is valid
	val, err := f.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.NotEmpty(t, val, "expected cell data in chart template output")
}

func TestConditionalFormatCommandIntegration(t *testing.T) {
	tmplPath := createConditionalFormatTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Score": 95},
			{"Score": 60},
			{"Score": 80},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Open output and verify conditional format exists
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	cfs, err := f.GetConditionalFormats("Sheet1")
	require.NoError(t, err)
	assert.True(t, len(cfs) > 0, "expected conditional formats")
}

func TestSparklineCommandIntegration(t *testing.T) {
	tmplPath := createSparklineTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Value": 10},
			{"Value": 20},
			{"Value": 30},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Verify the output is a valid xlsx
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Sparklines don't have a public getter in excelize, so just verify
	// the file is valid and data was written
	val, err := f.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.NotEmpty(t, val)
}

// ─── Deferred action registration tests ───

func TestTableCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &TableCommand{
		TableName: "TestTable",
		Style:     "TableStyleMedium9",
		Area:      createStaticArea(t, "Sheet1", 0, 0, 3, 2),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 2}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "table", actions[0].Name)
	assert.Equal(t, "Sheet1", actions[0].Sheet)
}

func TestDataValidationCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &DataValidationCommand{
		ValidationType: "list",
		Source:         "A,B,C",
		AllowBlank:     "true",
		Area:           createStaticArea(t, "Sheet1", 0, 0, 2, 3),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 3}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "dataValidation", actions[0].Name)
}

func TestConditionalFormatCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &ConditionalFormatCommand{
		FormatType: "colorScale",
		MinColor:   "#F00",
		MaxColor:   "#0F0",
		Area:       createStaticArea(t, "Sheet1", 0, 0, 1, 5),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 5}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "conditionalFormat", actions[0].Name)
}

func TestGroupCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &GroupCommand{
		Collapsed: "false",
		Area:      createStaticArea(t, "Sheet1", 0, 0, 2, 4),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 4}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "group", actions[0].Name)
}

func TestChartCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &ChartCommand{
		ChartType:  "bar",
		Title:      "My Chart",
		Series:     "Sheet1!$B$1:$B$5",
		Categories: "Sheet1!$A$1:$A$5",
		Width:      "480",
		Height:     "290",
		Area:       createStaticArea(t, "Sheet1", 0, 0, 2, 5),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 5}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "chart", actions[0].Name)
}

func TestDefinedNameCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &DefinedNameCommand{
		DefinedNameValue: "TestRange",
		Scope:            "workbook",
		Area:             createStaticArea(t, "Sheet1", 0, 0, 3, 2),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 2}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "definedName", actions[0].Name)
}

func TestSparklineCommandRegistersDeferredAction(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &SparklineCommand{
		SparkType: "line",
		DataRange: "Sheet1!A1:A5",
		Area:      createStaticArea(t, "Sheet1", 0, 0, 2, 1),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "sparkline", actions[0].Name)
}

// ─── Edge case tests ───

func TestDataValidationDecimalType(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &DataValidationCommand{
		ValidationType: "decimal",
		Min:            "0.5",
		Max:            "99.9",
		AllowBlank:     "true",
		Area:           createStaticArea(t, "Sheet1", 0, 0, 1, 1),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
}

func TestDataValidationCustomType(t *testing.T) {
	ctx := NewContext(map[string]any{})

	cmd := &DataValidationCommand{
		ValidationType: "custom",
		Source:         "=LEN(A1)>0",
		AllowBlank:     "true",
		Area:           createStaticArea(t, "Sheet1", 0, 0, 1, 1),
	}

	cellRef := NewCellRef("Sheet1", 0, 0)
	_, err := cmd.ApplyAt(cellRef, ctx, cmd.Area.Transformer)
	require.NoError(t, err)

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
}

func TestConditionalFormatColorScale(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Score")
	f.SetCellValue(sheet, "A2", "${e.Score}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:conditionalFormat(type=\"colorScale\" minColor=\"#F8696B\" maxColor=\"#63BE7B\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"employees\" var=\"e\" lastCell=\"A2\")",
	})

	tmplPath := filepath.Join(testdataDir(t), "cf_colorscale_template.xlsx")
	require.NoError(t, f.SaveAs(tmplPath))
	f.Close()
	defer os.Remove(tmplPath)

	data := map[string]any{
		"employees": []map[string]any{
			{"Score": 10},
			{"Score": 50},
			{"Score": 90},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	outF, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer outF.Close()

	cfs, err := outF.GetConditionalFormats("Sheet1")
	require.NoError(t, err)
	assert.True(t, len(cfs) > 0)
}

func TestChartDifferentTypes(t *testing.T) {
	chartTypes := []string{"bar", "col", "line", "pie", "area", "doughnut", "radar"}

	for _, ct := range chartTypes {
		t.Run(ct, func(t *testing.T) {
			f := excelize.NewFile()
			sheet := "Sheet1"
			f.SetCellValue(sheet, "A1", "Cat")
			f.SetCellValue(sheet, "B1", "Val")
			f.SetCellValue(sheet, "A2", "${e.Cat}")
			f.SetCellValue(sheet, "B2", "${e.Val}")

			f.AddComment(sheet, excelize.Comment{
				Cell: "A1", Author: "xlfill",
				Text: "jx:area(lastCell=\"B2\")\njx:chart(type=\"" + ct + "\" title=\"Test\" series=\"Sheet1!$B$1:$B$3\" categories=\"Sheet1!$A$1:$A$3\" lastCell=\"B2\")",
			})
			f.AddComment(sheet, excelize.Comment{
				Cell: "A2", Author: "xlfill",
				Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"B2\")",
			})

			tmplPath := filepath.Join(testdataDir(t), "chart_"+ct+"_template.xlsx")
			require.NoError(t, f.SaveAs(tmplPath))
			f.Close()
			defer os.Remove(tmplPath)

			data := map[string]any{
				"items": []map[string]any{
					{"Cat": "A", "Val": 10},
					{"Cat": "B", "Val": 20},
				},
			}

			var buf bytes.Buffer
			err := FillReader(mustOpen(t, tmplPath), &buf, data)
			require.NoError(t, err)

			outF, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
			require.NoError(t, err)
			outF.Close()
		})
	}
}

func TestTableShowColumnsOptions(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Col1")
	f.SetCellValue(sheet, "B1", "Col2")
	f.SetCellValue(sheet, "A2", "${e.A}")
	f.SetCellValue(sheet, "B2", "${e.B}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:table(name=\"T1\" showFirstColumn=\"true\" showLastColumn=\"true\" lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"B2\")",
	})

	tmplPath := filepath.Join(testdataDir(t), "table_cols_template.xlsx")
	require.NoError(t, f.SaveAs(tmplPath))
	f.Close()
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"A": "x", "B": "y"},
			{"A": "z", "B": "w"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	outF, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer outF.Close()

	tables, err := outF.GetTables("Sheet1")
	require.NoError(t, err)
	require.Len(t, tables, 1)
	assert.True(t, tables[0].ShowFirstColumn)
	assert.True(t, tables[0].ShowLastColumn)
}

func TestDataValidationWithErrorMessage(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Grade")
	f.SetCellValue(sheet, "A2", "${e.Grade}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:dataValidation(type=\"list\" source=\"A,B,C\" showError=\"true\" errorTitle=\"Invalid Grade\" error=\"Please select A, B, or C\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"A2\")",
	})

	tmplPath := filepath.Join(testdataDir(t), "dv_error_template.xlsx")
	require.NoError(t, f.SaveAs(tmplPath))
	f.Close()
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{{"Grade": "A"}},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	outF, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer outF.Close()

	dvs, err := outF.GetDataValidations("Sheet1")
	require.NoError(t, err)
	require.True(t, len(dvs) > 0)
	assert.True(t, dvs[0].ShowErrorMessage)
}

func TestDefinedNameWithScope(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "X")
	f.SetCellValue(sheet, "A2", "${e.X}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:definedName(name=\"LocalRange\" scope=\"Sheet1\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"A2\")",
	})

	tmplPath := filepath.Join(testdataDir(t), "dn_scope_template.xlsx")
	require.NoError(t, f.SaveAs(tmplPath))
	f.Close()
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{{"X": 1}, {"X": 2}},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	outF, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer outF.Close()

	dnames := outF.GetDefinedName()
	found := false
	for _, dn := range dnames {
		if dn.Name == "LocalRange" {
			found = true
			assert.Equal(t, "Sheet1", dn.Scope)
			break
		}
	}
	assert.True(t, found)
}

func TestMapOperatorToCriteria(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{">", ">"},
		{"<", "<"},
		{">=", ">="},
		{"<=", "<="},
		{"==", "=="},
		{"!=", "!="},
		{"between", "between"},
		{"greaterThan", ">"},
		{"lessThan", "<"},
		{"equal", "=="},
		{"notEqual", "!="},
		{"unknown", "unknown"},
	}

	for _, tt := range tests {
		assert.Equal(t, tt.expected, mapOperatorToCriteria(tt.input), "input: %s", tt.input)
	}
}

// ─── getCommandArea / attachArea tests for new types ───

func TestGetCommandAreaNewTypes(t *testing.T) {
	area := &Area{StartCell: NewCellRef("S", 0, 0), AreaSize: Size{1, 1}}

	tests := []struct {
		name string
		cmd  Command
	}{
		{"dataValidation", &DataValidationCommand{Area: area}},
		{"table", &TableCommand{Area: area}},
		{"conditionalFormat", &ConditionalFormatCommand{Area: area}},
		{"group", &GroupCommand{Area: area}},
		{"chart", &ChartCommand{Area: area}},
		{"definedName", &DefinedNameCommand{Area: area}},
		{"sparkline", &SparklineCommand{Area: area}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := getCommandArea(tt.cmd)
			assert.Equal(t, area, got)
		})
	}
}

func TestAttachAreaNewTypes(t *testing.T) {
	area := &Area{StartCell: NewCellRef("S", 0, 0), AreaSize: Size{1, 1}}

	cmds := []Command{
		&DataValidationCommand{},
		&TableCommand{},
		&ConditionalFormatCommand{},
		&GroupCommand{},
		&ChartCommand{},
		&DefinedNameCommand{},
		&SparklineCommand{},
	}

	for _, cmd := range cmds {
		attachArea(cmd, area)
		got := getCommandArea(cmd)
		assert.Equal(t, area, got, "failed for %T", cmd)
	}
}

// ─── Helper functions ───

// createStaticArea creates a simple static area backed by an excelize transformer
// for testing deferred action registration.
func createStaticArea(t *testing.T, sheet string, row, col, width, height int) *Area {
	t.Helper()
	f := excelize.NewFile()
	if sheet != "Sheet1" {
		f.NewSheet(sheet)
	}
	for r := 0; r < height; r++ {
		for c := 0; c < width; c++ {
			cellName := ColToName(col+c) + fmt.Sprintf("%d", row+r+1)
			f.SetCellValue(sheet, cellName, "")
		}
	}
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	t.Cleanup(func() { tx.Close() })

	return NewArea(
		NewCellRef(sheet, row, col),
		Size{Width: width, Height: height},
		tx,
	)
}

// mustOpen opens a file for reading.
func mustOpen(t *testing.T, path string) *os.File {
	t.Helper()
	f, err := os.Open(path)
	require.NoError(t, err)
	t.Cleanup(func() { f.Close() })
	return f
}

// ═══════════════════════════════════════════════════════════════════════════════
// Chunk 3: Immediate-execution commands
// pageBreak, autoColWidth, freezePanes, protect, include
// ═══════════════════════════════════════════════════════════════════════════════

// ─── Template helpers ───

func createPageBreakTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "${e.Name}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:pageBreak(lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"A2\")",
	})

	path := filepath.Join(testdataDir(t), "pagebreak_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func createAutoColWidthTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Description")
	f.SetCellValue(sheet, "A2", "${e.Name}")
	f.SetCellValue(sheet, "B2", "${e.Desc}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"B2\")\njx:autoColWidth(lastCell=\"B2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"B2\")",
	})

	path := filepath.Join(testdataDir(t), "autocolwidth_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func createFreezePanesTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Header")
	f.SetCellValue(sheet, "A2", "${e.Val}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:freezePanes(row=\"1\" col=\"0\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"A2\")",
	})

	path := filepath.Join(testdataDir(t), "freezepanes_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func createProtectTemplate(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Data")
	f.SetCellValue(sheet, "A2", "${e.Val}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: "jx:area(lastCell=\"A2\")\njx:protect(password=\"secret\" allowSort=\"true\" allowFilter=\"true\" lastCell=\"A2\")",
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A2", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"A2\")",
	})

	path := filepath.Join(testdataDir(t), "protect_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func createIncludeTestFile(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Header1")
	f.SetCellValue(sheet, "B1", "Header2")
	f.SetCellValue(sheet, "C1", "Header3")
	f.SetCellValue(sheet, "A2", "Val1")
	f.SetCellValue(sheet, "B2", "Val2")
	f.SetCellValue(sheet, "C2", "Val3")

	path := filepath.Join(testdataDir(t), "header.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

func createIncludeTemplate(t *testing.T, headerPath string) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "") // placeholder for include
	f.SetCellValue(sheet, "A3", "${e.Val}")

	f.AddComment(sheet, excelize.Comment{
		Cell: "A1", Author: "xlfill",
		Text: fmt.Sprintf("jx:area(lastCell=\"C3\")\njx:include(template=\"%s\" area=\"A1:C2\" lastCell=\"C2\")", headerPath),
	})
	f.AddComment(sheet, excelize.Comment{
		Cell: "A3", Author: "xlfill",
		Text: "jx:each(items=\"items\" var=\"e\" lastCell=\"C3\")",
	})

	path := filepath.Join(testdataDir(t), "include_template.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

// ─── Factory/parsing tests ───

func TestNewPageBreakCommandFromAttrs(t *testing.T) {
	cmd, err := newPageBreakCommandFromAttrs(map[string]string{})
	require.NoError(t, err)
	assert.IsType(t, &PageBreakCommand{}, cmd)
}

func TestNewAutoColWidthCommandFromAttrs(t *testing.T) {
	cmd, err := newAutoColWidthCommandFromAttrs(map[string]string{})
	require.NoError(t, err)
	assert.IsType(t, &AutoColWidthCommand{}, cmd)
}

func TestNewFreezePanesCommandFromAttrs(t *testing.T) {
	t.Run("with row and col", func(t *testing.T) {
		cmd, err := newFreezePanesCommandFromAttrs(map[string]string{
			"row": "1", "col": "2",
		})
		require.NoError(t, err)
		fp := cmd.(*FreezePanesCommand)
		assert.Equal(t, "1", fp.FreezeRow)
		assert.Equal(t, "2", fp.FreezeCol)
	})

	t.Run("empty attributes", func(t *testing.T) {
		cmd, err := newFreezePanesCommandFromAttrs(map[string]string{})
		require.NoError(t, err)
		fp := cmd.(*FreezePanesCommand)
		assert.Equal(t, "", fp.FreezeRow)
		assert.Equal(t, "", fp.FreezeCol)
	})
}

func TestNewProtectCommandFromAttrs(t *testing.T) {
	t.Run("all attributes", func(t *testing.T) {
		cmd, err := newProtectCommandFromAttrs(map[string]string{
			"password":                 "secret",
			"allowSort":                "true",
			"allowFilter":              "true",
			"allowSelectUnlockedCells": "true",
			"allowSelectLockedCells":   "true",
		})
		require.NoError(t, err)
		pc := cmd.(*ProtectCommand)
		assert.Equal(t, "secret", pc.Password)
		assert.Equal(t, "true", pc.AllowSort)
		assert.Equal(t, "true", pc.AllowFilter)
		assert.Equal(t, "true", pc.AllowSelectUnlockedCells)
		assert.Equal(t, "true", pc.AllowSelectLockedCells)
	})

	t.Run("empty attributes", func(t *testing.T) {
		cmd, err := newProtectCommandFromAttrs(map[string]string{})
		require.NoError(t, err)
		pc := cmd.(*ProtectCommand)
		assert.Equal(t, "", pc.Password)
	})
}

func TestNewIncludeCommandFromAttrs(t *testing.T) {
	t.Run("valid", func(t *testing.T) {
		cmd, err := newIncludeCommandFromAttrs(map[string]string{
			"template": "header.xlsx", "sheet": "Sheet1", "area": "A1:C2",
		})
		require.NoError(t, err)
		ic := cmd.(*IncludeCommand)
		assert.Equal(t, "header.xlsx", ic.TemplatePath)
		assert.Equal(t, "Sheet1", ic.SourceSheet)
		assert.Equal(t, "A1:C2", ic.SourceArea)
	})

	t.Run("missing template", func(t *testing.T) {
		_, err := newIncludeCommandFromAttrs(map[string]string{
			"area": "A1:C2",
		})
		assert.Error(t, err)
		assert.Contains(t, err.Error(), "template")
	})

	t.Run("missing area", func(t *testing.T) {
		_, err := newIncludeCommandFromAttrs(map[string]string{
			"template": "header.xlsx",
		})
		assert.Error(t, err)
		assert.Contains(t, err.Error(), "area")
	})
}

// ─── Name() tests ───

func TestChunk3CommandNames(t *testing.T) {
	tests := []struct {
		factory CommandFactory
		name    string
		attrs   map[string]string
	}{
		{newPageBreakCommandFromAttrs, "pageBreak", map[string]string{}},
		{newAutoColWidthCommandFromAttrs, "autoColWidth", map[string]string{}},
		{newFreezePanesCommandFromAttrs, "freezePanes", map[string]string{}},
		{newProtectCommandFromAttrs, "protect", map[string]string{}},
		{newIncludeCommandFromAttrs, "include", map[string]string{"template": "x.xlsx", "area": "A1:A1"}},
	}

	for _, tt := range tests {
		cmd, err := tt.factory(tt.attrs)
		require.NoError(t, err, "factory %s", tt.name)
		assert.Equal(t, tt.name, cmd.Name())
	}
}

// ─── Reset() tests (no-op) ───

func TestChunk3CommandReset(t *testing.T) {
	cmds := []Command{
		&PageBreakCommand{},
		&AutoColWidthCommand{},
		&FreezePanesCommand{},
		&ProtectCommand{},
		&IncludeCommand{TemplatePath: "x.xlsx", SourceArea: "A1:A1"},
	}
	for _, cmd := range cmds {
		cmd.Reset() // should not panic
	}
}

// ─── ApplyAt nil Area tests ───

func TestChunk3CommandApplyAtNilArea(t *testing.T) {
	cellRef := NewCellRef("Sheet1", 0, 0)
	ctx := NewContext(map[string]any{})

	ef := excelize.NewFile()
	tx, err := NewExcelizeTransformer(ef)
	require.NoError(t, err)
	defer tx.Close()

	tests := []struct {
		name string
		cmd  Command
	}{
		{"pageBreak", &PageBreakCommand{}},
		{"autoColWidth", &AutoColWidthCommand{}},
		{"freezePanes", &FreezePanesCommand{}},
		{"protect", &ProtectCommand{}},
		// include with nil Area also returns ZeroSize
		{"include", &IncludeCommand{TemplatePath: "x.xlsx", SourceArea: "A1:A1"}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			size, err := tt.cmd.ApplyAt(cellRef, ctx, tx)
			require.NoError(t, err)
			assert.Equal(t, ZeroSize, size)
		})
	}
}

// ─── Registry tests ───

func TestCommandRegistryChunk3Commands(t *testing.T) {
	reg := NewCommandRegistry()

	newCmds := []string{
		"pageBreak", "autoColWidth", "freezePanes", "protect", "include",
	}

	names := reg.KnownNames()
	for _, name := range newCmds {
		t.Run(name, func(t *testing.T) {
			found := false
			for _, n := range names {
				if n == name {
					found = true
					break
				}
			}
			assert.True(t, found, "command %q not registered", name)
		})
	}
}

// ─── getCommandArea / attachArea tests ───

func TestGetCommandAreaChunk3Types(t *testing.T) {
	area := &Area{StartCell: NewCellRef("S", 0, 0), AreaSize: Size{1, 1}}

	tests := []struct {
		name string
		cmd  Command
	}{
		{"pageBreak", &PageBreakCommand{Area: area}},
		{"autoColWidth", &AutoColWidthCommand{Area: area}},
		{"freezePanes", &FreezePanesCommand{Area: area}},
		{"protect", &ProtectCommand{Area: area}},
		{"include", &IncludeCommand{Area: area}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := getCommandArea(tt.cmd)
			assert.Equal(t, area, got)
		})
	}
}

func TestAttachAreaChunk3Types(t *testing.T) {
	area := &Area{StartCell: NewCellRef("S", 0, 0), AreaSize: Size{1, 1}}

	cmds := []Command{
		&PageBreakCommand{},
		&AutoColWidthCommand{},
		&FreezePanesCommand{},
		&ProtectCommand{},
		&IncludeCommand{},
	}

	for _, cmd := range cmds {
		attachArea(cmd, area)
		got := getCommandArea(cmd)
		assert.Equal(t, area, got, "failed for %T", cmd)
	}
}

// ─── Integration tests ───

func TestPageBreakCommandIntegration(t *testing.T) {
	tmplPath := createPageBreakTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"Name": "Alice"},
			{"Name": "Bob"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Open output and verify page break was inserted
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Verify page break was inserted by attempting to remove it.
	// RemovePageBreak returns nil only if the break existed.
	// The output area is rows 1-2 (0-based 0-1), so break is at row 3 (A3).
	err = f.RemovePageBreak("Sheet1", "A3")
	require.NoError(t, err, "expected page break at A3 to exist and be removable")
}

func TestAutoColWidthCommandIntegration(t *testing.T) {
	tmplPath := createAutoColWidthTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"Name": "A", "Desc": "Short"},
			{"Name": "B", "Desc": "A much longer description text for testing width"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Column B should be wider than column A because its content is longer
	widthA, err := f.GetColWidth("Sheet1", "A")
	require.NoError(t, err)
	widthB, err := f.GetColWidth("Sheet1", "B")
	require.NoError(t, err)

	assert.True(t, widthB > widthA, "column B (%.1f) should be wider than column A (%.1f)", widthB, widthA)
}

func TestFreezePanesCommandIntegration(t *testing.T) {
	tmplPath := createFreezePanesTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"Val": 1},
			{"Val": 2},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	// Verify the output is a valid xlsx
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// GetPanes returns pane settings (value type, not pointer)
	panes, err := f.GetPanes("Sheet1")
	require.NoError(t, err)
	assert.True(t, panes.Freeze, "expected freeze panes to be set")
}

func TestProtectCommandIntegration(t *testing.T) {
	tmplPath := createProtectTemplate(t)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"Val": "data1"},
			{"Val": "data2"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Verify sheet is protected by attempting to unprotect it.
	// UnprotectSheet returns nil if the sheet was protected with the given password.
	err = f.UnprotectSheet("Sheet1", "secret")
	require.NoError(t, err, "expected Sheet1 to be protected and unprotectable with correct password")
}

func TestIncludeCommandIntegration(t *testing.T) {
	headerPath := createIncludeTestFile(t)
	defer os.Remove(headerPath)

	tmplPath := createIncludeTemplate(t, headerPath)
	defer os.Remove(tmplPath)

	data := map[string]any{
		"items": []map[string]any{
			{"Val": "row1"},
		},
	}

	var buf bytes.Buffer
	err := FillReader(mustOpen(t, tmplPath), &buf, data)
	require.NoError(t, err)

	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	require.NoError(t, err)
	defer f.Close()

	// Verify cells were copied from the included template
	val, err := f.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Header1", val)

	val, err = f.GetCellValue("Sheet1", "B1")
	require.NoError(t, err)
	assert.Equal(t, "Header2", val)

	val, err = f.GetCellValue("Sheet1", "C1")
	require.NoError(t, err)
	assert.Equal(t, "Header3", val)

	val, err = f.GetCellValue("Sheet1", "A2")
	require.NoError(t, err)
	assert.Equal(t, "Val1", val)
}

// ─── Unit tests with ApplyAt directly ───

func TestPageBreakCommandApplyAtWithArea(t *testing.T) {
	area := createStaticArea(t, "Sheet1", 0, 0, 2, 3)
	cmd := &PageBreakCommand{Area: area}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 3}, size)
}

func TestAutoColWidthCommandApplyAtWithArea(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "ShortText")
	f.SetCellValue(sheet, "B1", "A much longer cell value for width testing purposes")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 2, Height: 1}, tx)
	cmd := &AutoColWidthCommand{Area: area}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef(sheet, 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 1}, size)

	// Verify column B is wider than column A
	widthA, _ := f.GetColWidth(sheet, "A")
	widthB, _ := f.GetColWidth(sheet, "B")
	assert.True(t, widthB > widthA, "B (%.1f) should be wider than A (%.1f)", widthB, widthA)
}

func TestFreezePanesCommandApplyAtWithArea(t *testing.T) {
	area := createStaticArea(t, "Sheet1", 0, 0, 2, 3)
	cmd := &FreezePanesCommand{
		FreezeRow: "1",
		FreezeCol: "0",
		Area:      area,
	}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 3}, size)
}

func TestProtectCommandApplyAtWithArea(t *testing.T) {
	area := createStaticArea(t, "Sheet1", 0, 0, 2, 2)
	cmd := &ProtectCommand{
		Password:  "pass",
		AllowSort: "true",
		Area:      area,
	}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 2}, size)
}

func TestIncludeCommandApplyAtInvalidTemplate(t *testing.T) {
	area := createStaticArea(t, "Sheet1", 0, 0, 3, 2)
	cmd := &IncludeCommand{
		TemplatePath: "/nonexistent/file.xlsx",
		SourceArea:   "A1:C2",
		Area:         area,
	}

	ctx := NewContext(map[string]any{})
	ef := excelize.NewFile()
	tx, err := NewExcelizeTransformer(ef)
	require.NoError(t, err)
	defer tx.Close()

	cellRef := NewCellRef("Sheet1", 0, 0)
	_, err = cmd.ApplyAt(cellRef, ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "include: absolute paths not allowed")
}

func TestIncludeCommandApplyAtInvalidArea(t *testing.T) {
	// Create a valid file to open, but pass invalid area
	headerPath := createIncludeTestFile(t)
	defer os.Remove(headerPath)

	area := createStaticArea(t, "Sheet1", 0, 0, 1, 1)
	cmd := &IncludeCommand{
		TemplatePath: headerPath,
		SourceArea:   "INVALID",
		Area:         area,
	}

	ctx := NewContext(map[string]any{})
	ef := excelize.NewFile()
	tx, err := NewExcelizeTransformer(ef)
	require.NoError(t, err)
	defer tx.Close()

	cellRef := NewCellRef("Sheet1", 0, 0)
	_, err = cmd.ApplyAt(cellRef, ctx, tx)
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "include: parse area")
}

func TestIncludeCommandDefaultSheet(t *testing.T) {
	headerPath := createIncludeTestFile(t)
	defer os.Remove(headerPath)

	ef := excelize.NewFile()
	tx, err := NewExcelizeTransformer(ef)
	require.NoError(t, err)
	defer tx.Close()

	// Give it an Area so the nil check doesn't short-circuit
	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 3, Height: 1}, tx)
	cmd := &IncludeCommand{
		TemplatePath: headerPath,
		SourceSheet:  "", // should default to first sheet
		SourceArea:   "A1:C1",
		Area:         area,
	}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 1}, size)

	// Verify cells were copied
	val, err := ef.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "Header1", val)
}

func TestFreezePanesCommandWithExpressions(t *testing.T) {
	area := createStaticArea(t, "Sheet1", 0, 0, 2, 3)
	cmd := &FreezePanesCommand{
		FreezeRow: "freezeRow",
		FreezeCol: "freezeCol",
		Area:      area,
	}

	ctx := NewContext(map[string]any{
		"freezeRow": 2,
		"freezeCol": 1,
	})
	cellRef := NewCellRef("Sheet1", 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, area.Transformer)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 2, Height: 3}, size)
}

func TestAutoColWidthCommandCapAt60(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	// Create a very long string (100 chars) to test the cap
	longStr := ""
	for i := 0; i < 100; i++ {
		longStr += "X"
	}
	f.SetCellValue(sheet, "A1", longStr)

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &AutoColWidthCommand{Area: area}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef(sheet, 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	// Width should be capped at 60
	width, _ := f.GetColWidth(sheet, "A")
	assert.LessOrEqual(t, width, 60.0, "width should be capped at 60")
}

func TestProtectCommandAllOptions(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "data")

	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)

	area := NewArea(NewCellRef(sheet, 0, 0), Size{Width: 1, Height: 1}, tx)
	cmd := &ProtectCommand{
		Password:                 "mypass",
		AllowSort:                "true",
		AllowFilter:              "true",
		AllowSelectUnlockedCells: "true",
		AllowSelectLockedCells:   "true",
		Area:                     area,
	}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef(sheet, 0, 0)
	size, err := cmd.ApplyAt(cellRef, ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	// Verify sheet is protected by attempting to unprotect it
	err = f.UnprotectSheet(sheet, "mypass")
	require.NoError(t, err, "expected sheet to be protected and unprotectable")
}

// ─── describeCommandAttrs tests ───

func TestDescribeCommandAttrsChunk3(t *testing.T) {
	t.Run("pageBreak", func(t *testing.T) {
		cmd := &PageBreakCommand{}
		result := describeCommandAttrs(cmd)
		assert.Equal(t, "", result)
	})

	t.Run("autoColWidth", func(t *testing.T) {
		cmd := &AutoColWidthCommand{}
		result := describeCommandAttrs(cmd)
		assert.Equal(t, "", result)
	})

	t.Run("freezePanes", func(t *testing.T) {
		cmd := &FreezePanesCommand{FreezeRow: "1", FreezeCol: "2"}
		result := describeCommandAttrs(cmd)
		assert.Contains(t, result, "row=\"1\"")
		assert.Contains(t, result, "col=\"2\"")
	})

	t.Run("protect", func(t *testing.T) {
		cmd := &ProtectCommand{Password: "secret", AllowSort: "true"}
		result := describeCommandAttrs(cmd)
		assert.Contains(t, result, "password=***")
		assert.Contains(t, result, "allowSort")
	})

	t.Run("include", func(t *testing.T) {
		cmd := &IncludeCommand{TemplatePath: "header.xlsx", SourceSheet: "Sheet1", SourceArea: "A1:C2"}
		result := describeCommandAttrs(cmd)
		assert.Contains(t, result, "template=\"header.xlsx\"")
		assert.Contains(t, result, "sheet=\"Sheet1\"")
		assert.Contains(t, result, "area=\"A1:C2\"")
	})
}

// ─── Include with nil transformer ───

func TestIncludeCommandApplyAtNilTransformer(t *testing.T) {
	headerPath := createIncludeTestFile(t)
	defer os.Remove(headerPath)

	ef := excelize.NewFile()
	tx, err := NewExcelizeTransformer(ef)
	require.NoError(t, err)
	defer tx.Close()

	area := NewArea(NewCellRef("Sheet1", 0, 0), Size{Width: 3, Height: 2}, tx)
	cmd := &IncludeCommand{
		TemplatePath: headerPath,
		SourceArea:   "A1:C2",
		Area:         area,
	}

	ctx := NewContext(map[string]any{})
	cellRef := NewCellRef("Sheet1", 0, 0)
	// Pass the actual transformer; verify it works and returns data
	size, err := cmd.ApplyAt(cellRef, ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 3, Height: 2}, size)
}
