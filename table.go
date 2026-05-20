package xlfill

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// TableCommand implements the jx:table command.
// It processes its inner area normally, then registers a deferred action
// that creates an Excel table on the expanded output range.
type TableCommand struct {
	TableName       string // table name
	Style           string // table style (default "TableStyleMedium9")
	ShowFirstColumn string // "true" or "false"
	ShowLastColumn  string // "true" or "false"
	Area            *Area
}

func (c *TableCommand) Name() string    { return "table" }
func (c *TableCommand) Reset()          {}
func (c *TableCommand) GetArea() *Area  { return c.Area }
func (c *TableCommand) SetArea(a *Area) { c.Area = a }

// newTableCommandFromAttrs creates a TableCommand from parsed attributes.
func newTableCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &TableCommand{
		TableName:       attrs["name"],
		Style:           attrs["style"],
		ShowFirstColumn: attrs["showFirstColumn"],
		ShowLastColumn:  attrs["showLastColumn"],
	}
	if cmd.TableName == "" {
		return nil, fmt.Errorf("table command requires 'name' attribute")
	}
	if cmd.Style == "" {
		cmd.Style = "TableStyleMedium9"
	}
	return cmd, nil
}

// ApplyAt processes the inner area, then registers a deferred action for table creation.
func (c *TableCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	// Process the inner area normally first
	size, err := c.Area.ApplyAt(cellRef, ctx)
	if err != nil {
		return ZeroSize, err
	}

	// Capture values for the closure
	sheet := cellRef.Sheet
	startRow := cellRef.Row
	startCol := cellRef.Col
	endRow := cellRef.Row + size.Height - 1
	endCol := cellRef.Col + size.Width - 1

	tableName := c.TableName
	style := c.Style
	showFirstColumn := c.ShowFirstColumn == "true"
	showLastColumn := c.ShowLastColumn == "true"

	// Register deferred action
	if err := ctx.RegisterDeferred(DeferredAction{
		Name:     "table",
		Sheet:    sheet,
		StartRow: startRow,
		StartCol: startCol,
		EndRow:   endRow,
		EndCol:   endCol,
		Execute: func(tx *ExcelizeTransformer) error {
			startCellName := NewCellRef(sheet, startRow, startCol).CellName()
			endCellName := NewCellRef(sheet, endRow, endCol).CellName()
			rangeStr := fmt.Sprintf("%s:%s", startCellName, endCellName)

			return tx.file.AddTable(sheet, &excelize.Table{
				Range:           rangeStr,
				Name:            tableName,
				StyleName:       style,
				ShowFirstColumn: showFirstColumn,
				ShowLastColumn:  showLastColumn,
			})
		},
	}); err != nil {
		return ZeroSize, fmt.Errorf("table: %w", err)
	}

	return size, nil
}
