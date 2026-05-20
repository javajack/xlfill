package xlfill

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// FreezePanesCommand implements the jx:freezePanes command.
// It processes its inner area, then sets a freeze pane split on the sheet.
type FreezePanesCommand struct {
	FreezeRow string // 0-based row to freeze below (expression or literal)
	FreezeCol string // 0-based column to freeze right of (expression or literal)
	Area      *Area
}

func (c *FreezePanesCommand) Name() string    { return "freezePanes" }
func (c *FreezePanesCommand) Reset()          {}
func (c *FreezePanesCommand) GetArea() *Area  { return c.Area }
func (c *FreezePanesCommand) SetArea(a *Area) { c.Area = a }

// newFreezePanesCommandFromAttrs creates a FreezePanesCommand from parsed attributes.
func newFreezePanesCommandFromAttrs(attrs map[string]string) (Command, error) {
	return &FreezePanesCommand{
		FreezeRow: attrs["row"],
		FreezeCol: attrs["col"],
	}, nil
}

// ApplyAt processes the inner area, then sets a freeze pane on the sheet.
func (c *FreezePanesCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	size, err := c.Area.ApplyAt(cellRef, ctx)
	if err != nil {
		return ZeroSize, err
	}

	etx := unwrapExcelizeTransformer(transformer)
	if etx == nil {
		return size, nil
	}

	// Parse row/col - can be expressions or literals
	freezeRow := 0
	if c.FreezeRow != "" {
		val, evalErr := ctx.Evaluate(c.FreezeRow)
		if evalErr == nil {
			freezeRow = toInt(val)
		} else {
			// Try direct parse
			if n, parseErr := strconv.Atoi(c.FreezeRow); parseErr == nil {
				freezeRow = n
			}
		}
	}

	freezeCol := 0
	if c.FreezeCol != "" {
		val, evalErr := ctx.Evaluate(c.FreezeCol)
		if evalErr == nil {
			freezeCol = toInt(val)
		} else {
			if n, parseErr := strconv.Atoi(c.FreezeCol); parseErr == nil {
				freezeCol = n
			}
		}
	}

	// excelize SetPanes API
	freezeCell := ColToName(freezeCol) + strconv.Itoa(freezeRow+1)
	if err := etx.file.SetPanes(cellRef.Sheet, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      freezeCol,
		YSplit:      freezeRow,
		TopLeftCell: freezeCell,
		ActivePane:  "bottomRight",
	}); err != nil {
		return size, fmt.Errorf("set freeze panes: %w", err)
	}

	return size, nil
}
