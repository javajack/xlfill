package goxls

import "fmt"

// CellDataUpdater is an interface for custom cell processing.
// Users implement this to modify cell data during template processing.
type CellDataUpdater interface {
	UpdateCellData(cellData *CellData, targetCell CellRef, ctx *Context)
}

// UpdateCellCommand implements the jx:updateCell command.
// It delegates cell modification to a CellDataUpdater from the context.
type UpdateCellCommand struct {
	Updater string // context key for CellDataUpdater
	Area    *Area
}

func (c *UpdateCellCommand) Name() string { return "updateCell" }
func (c *UpdateCellCommand) Reset()       {}

// newUpdateCellCommandFromAttrs creates an UpdateCellCommand from parsed attributes.
func newUpdateCellCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &UpdateCellCommand{
		Updater: attrs["updater"],
	}
	if cmd.Updater == "" {
		return nil, fmt.Errorf("updateCell command requires 'updater' attribute")
	}
	return cmd, nil
}

// ApplyAt applies the cell updater at the target position.
func (c *UpdateCellCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	// Look up updater from context
	updaterVal := ctx.GetVar(c.Updater)
	if updaterVal == nil {
		return ZeroSize, fmt.Errorf("updater %q not found in context", c.Updater)
	}

	updater, ok := updaterVal.(CellDataUpdater)
	if !ok {
		return ZeroSize, fmt.Errorf("context variable %q does not implement CellDataUpdater", c.Updater)
	}

	// First transform the area normally
	if c.Area != nil {
		size, err := c.Area.ApplyAt(cellRef, ctx)
		if err != nil {
			return ZeroSize, err
		}

		// Then apply updater to each target cell
		for row := 0; row < size.Height; row++ {
			for col := 0; col < size.Width; col++ {
				targetRef := NewCellRef(cellRef.Sheet, cellRef.Row+row, cellRef.Col+col)
				cd := transformer.GetCellData(targetRef)
				if cd == nil {
					cd = &CellData{Ref: targetRef}
				}
				updater.UpdateCellData(cd, targetRef, ctx)

				// Apply updated value
				if cd.Formula != "" {
					transformer.SetFormula(targetRef, cd.Formula)
				} else if cd.Value != nil {
					transformer.SetCellValue(targetRef, cd.Value)
				}
			}
		}
		return size, nil
	}

	// No area — just apply to the single cell
	cd := transformer.GetCellData(cellRef)
	if cd == nil {
		cd = &CellData{Ref: cellRef}
	}
	updater.UpdateCellData(cd, cellRef, ctx)

	if cd.Formula != "" {
		transformer.SetFormula(cellRef, cd.Formula)
	} else if cd.Value != nil {
		transformer.SetCellValue(cellRef, cd.Value)
	}

	return Size{Width: 1, Height: 1}, nil
}
