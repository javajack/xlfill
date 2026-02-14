package xlfill

import "fmt"

// CommandBinding binds a Command to the area it operates on within a parent area.
type CommandBinding struct {
	Command  Command
	StartRef CellRef // start cell of this command's area (relative to parent)
	Size     Size    // size of this command's area
}

// Area represents a rectangular region in a worksheet that can be processed.
type Area struct {
	StartCell   CellRef
	AreaSize    Size
	Bindings    []*CommandBinding
	Transformer Transformer
	Listeners   []AreaListener
}

// NewArea creates a new Area.
func NewArea(start CellRef, size Size, transformer Transformer) *Area {
	return &Area{
		StartCell:   start,
		AreaSize:    size,
		Transformer: transformer,
	}
}

// AddCommand adds a command binding to this area.
func (a *Area) AddCommand(cmd Command, startRef CellRef, size Size) {
	a.Bindings = append(a.Bindings, &CommandBinding{
		Command:  cmd,
		StartRef: startRef,
		Size:     size,
	})
}

// ApplyAt processes this area at the given target cell position.
// It transforms all cells, executing embedded commands as encountered.
func (a *Area) ApplyAt(targetCell CellRef, ctx *Context) (Size, error) {
	if a.Transformer == nil {
		return ZeroSize, fmt.Errorf("area has no transformer")
	}

	// If no commands, just transform all cells (static area)
	if len(a.Bindings) == 0 {
		return a.transformStaticArea(targetCell, ctx)
	}

	// Process with commands
	return a.processWithCommands(targetCell, ctx)
}

// transformStaticArea transforms all cells in the area without any command processing.
func (a *Area) transformStaticArea(targetCell CellRef, ctx *Context) (Size, error) {
	for row := 0; row < a.AreaSize.Height; row++ {
		for col := 0; col < a.AreaSize.Width; col++ {
			srcRef := NewCellRef(a.StartCell.Sheet, a.StartCell.Row+row, a.StartCell.Col+col)
			dstRef := NewCellRef(targetCell.Sheet, targetCell.Row+row, targetCell.Col+col)
			if err := a.transformCell(srcRef, dstRef, ctx); err != nil {
				return ZeroSize, fmt.Errorf("transform cell %s → %s: %w", srcRef, dstRef, err)
			}
		}
	}
	return a.AreaSize, nil
}

// transformCell transforms a single cell, firing listeners and injecting built-in variables.
func (a *Area) transformCell(src, target CellRef, ctx *Context) error {
	// Inject built-in position variables
	ctx.setRunVar("_row", target.Row+1) // 1-based row number
	ctx.setRunVar("_col", target.Col)   // 0-based column index

	// Fire before-transform listeners
	for _, l := range a.Listeners {
		if !l.BeforeTransformCell(src, target, ctx, a.Transformer) {
			// Listener says skip default transform
			for _, l2 := range a.Listeners {
				l2.AfterTransformCell(src, target, ctx, a.Transformer)
			}
			return nil
		}
	}

	if err := a.Transformer.Transform(src, target, ctx, true); err != nil {
		return err
	}

	// Fire after-transform listeners
	for _, l := range a.Listeners {
		l.AfterTransformCell(src, target, ctx, a.Transformer)
	}
	return nil
}

// processWithCommands processes the area, executing commands and transforming static cells.
// Commands may occupy a sub-region of a row. Static cells on the same row but outside
// the command's column range are transformed alongside the command.
func (a *Area) processWithCommands(targetCell CellRef, ctx *Context) (Size, error) {
	totalHeight := 0
	maxWidth := a.AreaSize.Width
	currentTargetRow := targetCell.Row

	prevCmdEndRow := a.StartCell.Row // tracks where we are in source

	for _, binding := range a.Bindings {
		cmdSrcStartRow := binding.StartRef.Row

		// Transform static rows between previous command end and this command start
		staticRows := cmdSrcStartRow - prevCmdEndRow
		if staticRows > 0 {
			if err := a.transformRows(prevCmdEndRow, staticRows, targetCell.Sheet, currentTargetRow, targetCell.Col, ctx, nil); err != nil {
				return ZeroSize, err
			}
			currentTargetRow += staticRows
			totalHeight += staticRows
		}

		// Transform static cells on the command's row(s) that are outside the command's column range
		cmdColStart := binding.StartRef.Col - a.StartCell.Col // relative col offset
		cmdColEnd := cmdColStart + binding.Size.Width
		cmdRowCount := binding.Size.Height

		if err := a.transformRows(binding.StartRef.Row, cmdRowCount, targetCell.Sheet, currentTargetRow, targetCell.Col, ctx, &colExclusion{start: cmdColStart, end: cmdColEnd}); err != nil {
			return ZeroSize, err
		}

		// Execute command
		cmdTarget := NewCellRef(targetCell.Sheet, currentTargetRow, targetCell.Col+cmdColStart)
		cmdSize, err := binding.Command.ApplyAt(cmdTarget, ctx, a.Transformer)
		if err != nil {
			return ZeroSize, fmt.Errorf("command %s at %s: %w", binding.Command.Name(), cmdTarget, err)
		}

		// Determine how many target rows this command band occupies.
		// If the command spans the full area width, use command's actual height (allows contraction).
		// If it's a partial-width command (static cells share the row), use at least source height.
		rowsConsumed := cmdSize.Height
		hasStaticCols := cmdColStart > 0 || cmdColEnd < a.AreaSize.Width
		if hasStaticCols && rowsConsumed < cmdRowCount {
			rowsConsumed = cmdRowCount
		}
		currentTargetRow += rowsConsumed
		totalHeight += rowsConsumed
		if cmdSize.Width+cmdColStart > maxWidth {
			maxWidth = cmdSize.Width + cmdColStart
		}

		prevCmdEndRow = binding.StartRef.Row + binding.Size.Height
	}

	// Transform static rows after last command
	remainingRows := (a.StartCell.Row + a.AreaSize.Height) - prevCmdEndRow
	if remainingRows > 0 {
		if err := a.transformRows(prevCmdEndRow, remainingRows, targetCell.Sheet, currentTargetRow, targetCell.Col, ctx, nil); err != nil {
			return ZeroSize, err
		}
		totalHeight += remainingRows
	}

	return Size{Width: maxWidth, Height: totalHeight}, nil
}

// colExclusion defines a column range to skip during row transformation.
type colExclusion struct {
	start int // inclusive, relative to area
	end   int // exclusive, relative to area
}

// transformRows transforms rows from the source area to target, optionally excluding a column range.
func (a *Area) transformRows(srcStartRow, rowCount int, targetSheet string, targetStartRow, targetStartCol int, ctx *Context, exclude *colExclusion) error {
	for row := 0; row < rowCount; row++ {
		srcRow := srcStartRow + row
		for col := 0; col < a.AreaSize.Width; col++ {
			if exclude != nil && col >= exclude.start && col < exclude.end {
				continue
			}
			srcRef := NewCellRef(a.StartCell.Sheet, srcRow, a.StartCell.Col+col)
			dstRef := NewCellRef(targetSheet, targetStartRow+row, targetStartCol+col)
			if err := a.transformCell(srcRef, dstRef, ctx); err != nil {
				return err
			}
		}
	}
	return nil
}

// ClearCells clears all template cells in this area.
func (a *Area) ClearCells() {
	if a.Transformer == nil {
		return
	}
	for row := 0; row < a.AreaSize.Height; row++ {
		for col := 0; col < a.AreaSize.Width; col++ {
			ref := NewCellRef(a.StartCell.Sheet, a.StartCell.Row+row, a.StartCell.Col+col)
			a.Transformer.ClearCell(ref)
		}
	}
}
