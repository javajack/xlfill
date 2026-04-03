package xlfill

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// DefinedNameCommand implements the jx:definedName command.
// It processes its inner area normally, then registers a deferred action
// that creates an Excel defined name (named range) for the expanded output range.
type DefinedNameCommand struct {
	DefinedNameValue string // the defined name
	Scope            string // scope: "workbook" (default) or a sheet name
	Area             *Area
}

func (c *DefinedNameCommand) Name() string { return "definedName" }
func (c *DefinedNameCommand) Reset()       {}

// newDefinedNameCommandFromAttrs creates a DefinedNameCommand from parsed attributes.
func newDefinedNameCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &DefinedNameCommand{
		DefinedNameValue: attrs["name"],
		Scope:            attrs["scope"],
	}
	if cmd.DefinedNameValue == "" {
		return nil, fmt.Errorf("definedName command requires 'name' attribute")
	}
	if cmd.Scope == "" {
		cmd.Scope = "workbook"
	}
	return cmd, nil
}

// ApplyAt processes the inner area, then registers a deferred action for the defined name.
func (c *DefinedNameCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
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

	definedName := c.DefinedNameValue
	scope := c.Scope

	// Register deferred action
	ctx.RegisterDeferred(DeferredAction{
		Name:     "definedName",
		Sheet:    sheet,
		StartRow: startRow,
		StartCol: startCol,
		EndRow:   endRow,
		EndCol:   endCol,
		Execute: func(tx *ExcelizeTransformer) error {
			startColName := ColToName(startCol)
			endColName := ColToName(endCol)
			refersTo := fmt.Sprintf("'%s'!$%s$%d:$%s$%d", sheet, startColName, startRow+1, endColName, endRow+1)

			dn := &excelize.DefinedName{
				Name:     definedName,
				RefersTo: refersTo,
				Scope:    scope,
			}

			return tx.file.SetDefinedName(dn)
		},
	})

	return size, nil
}
