package goxls

import (
	"fmt"
	"reflect"
	"strings"
)

// GridCommand implements the jx:grid command for dynamic grid rendering.
// It renders headers horizontally and data rows below.
type GridCommand struct {
	Headers    string // expression for header values ([]any)
	Data       string // expression for data rows ([]any)
	Props      string // comma-separated property names for object data
	FormatCells string // type-to-format mapping (unused for now)
	HeaderArea *Area
	BodyArea   *Area
}

func (c *GridCommand) Name() string { return "grid" }
func (c *GridCommand) Reset()       {}

// newGridCommandFromAttrs creates a GridCommand from parsed attributes.
func newGridCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &GridCommand{
		Headers:    attrs["headers"],
		Data:       attrs["data"],
		Props:      attrs["props"],
		FormatCells: attrs["formatCells"],
	}
	if cmd.Headers == "" {
		return nil, fmt.Errorf("grid command requires 'headers' attribute")
	}
	if cmd.Data == "" {
		return nil, fmt.Errorf("grid command requires 'data' attribute")
	}
	return cmd, nil
}

// ApplyAt renders the grid at the given target cell.
func (c *GridCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	// Evaluate headers
	headersVal, err := ctx.Evaluate(c.Headers)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate headers %q: %w", c.Headers, err)
	}
	headers, err := toSlice(headersVal)
	if err != nil {
		return ZeroSize, fmt.Errorf("headers not iterable: %w", err)
	}

	// Evaluate data
	dataVal, err := ctx.Evaluate(c.Data)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate data %q: %w", c.Data, err)
	}
	dataRows, err := toSlice(dataVal)
	if err != nil {
		return ZeroSize, fmt.Errorf("data not iterable: %w", err)
	}

	if len(headers) == 0 {
		return ZeroSize, nil
	}

	totalWidth := len(headers)
	totalHeight := 0

	// Render headers (one per column)
	for col, header := range headers {
		target := NewCellRef(cellRef.Sheet, cellRef.Row, cellRef.Col+col)
		transformer.SetCellValue(target, header)
	}
	totalHeight++ // header row

	// Parse props if provided
	var propNames []string
	if c.Props != "" {
		for _, p := range strings.Split(c.Props, ",") {
			propNames = append(propNames, strings.TrimSpace(p))
		}
	}

	// Render data rows
	for rowIdx, row := range dataRows {
		rowSlice, err := extractRowData(row, propNames)
		if err != nil {
			return ZeroSize, fmt.Errorf("extract row %d data: %w", rowIdx, err)
		}
		for col := 0; col < totalWidth && col < len(rowSlice); col++ {
			target := NewCellRef(cellRef.Sheet, cellRef.Row+1+rowIdx, cellRef.Col+col)
			transformer.SetCellValue(target, rowSlice[col])
		}
		totalHeight++
	}

	return Size{Width: totalWidth, Height: totalHeight}, nil
}

// extractRowData extracts values from a data row.
func extractRowData(row any, propNames []string) ([]any, error) {
	if row == nil {
		return nil, nil
	}

	// If it's already a slice, use directly
	if slice, err := toSlice(row); err == nil {
		return slice, nil
	}

	// If propNames specified, extract those properties
	if len(propNames) > 0 {
		result := make([]any, len(propNames))
		for i, prop := range propNames {
			result[i] = getField(row, prop)
		}
		return result, nil
	}

	// Try to extract all fields from struct/map
	v := reflect.ValueOf(row)
	if v.Kind() == reflect.Map {
		result := make([]any, 0, v.Len())
		for _, key := range v.MapKeys() {
			result = append(result, v.MapIndex(key).Interface())
		}
		return result, nil
	}

	return []any{row}, nil
}
