package goxls

import (
	"fmt"
	"strconv"
)

// MergeCellsCommand implements the jx:mergeCells command.
type MergeCellsCommand struct {
	Cols    string // number of columns to merge (expression)
	Rows    string // number of rows to merge (expression)
	MinCols string // minimum cols before merging
	MinRows string // minimum rows before merging
}

func (c *MergeCellsCommand) Name() string { return "mergeCells" }
func (c *MergeCellsCommand) Reset()       {}

// newMergeCellsCommandFromAttrs creates a MergeCellsCommand from parsed attributes.
func newMergeCellsCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &MergeCellsCommand{
		Cols:    attrs["cols"],
		Rows:    attrs["rows"],
		MinCols: attrs["minCols"],
		MinRows: attrs["minRows"],
	}
	return cmd, nil
}

// ApplyAt merges cells at the target position.
func (c *MergeCellsCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	cols := 1
	rows := 1

	if c.Cols != "" {
		val, err := ctx.Evaluate(c.Cols)
		if err != nil {
			// Try direct integer parse
			if n, parseErr := strconv.Atoi(c.Cols); parseErr == nil {
				cols = n
			} else {
				return ZeroSize, fmt.Errorf("evaluate cols %q: %w", c.Cols, err)
			}
		} else {
			cols = toInt(val)
		}
	}

	if c.Rows != "" {
		val, err := ctx.Evaluate(c.Rows)
		if err != nil {
			if n, parseErr := strconv.Atoi(c.Rows); parseErr == nil {
				rows = n
			} else {
				return ZeroSize, fmt.Errorf("evaluate rows %q: %w", c.Rows, err)
			}
		} else {
			rows = toInt(val)
		}
	}

	// Check minimum thresholds
	if c.MinCols != "" {
		minCols := 0
		if n, err := strconv.Atoi(c.MinCols); err == nil {
			minCols = n
		}
		if cols < minCols {
			return Size{Width: cols, Height: rows}, nil // skip merge
		}
	}
	if c.MinRows != "" {
		minRows := 0
		if n, err := strconv.Atoi(c.MinRows); err == nil {
			minRows = n
		}
		if rows < minRows {
			return Size{Width: cols, Height: rows}, nil // skip merge
		}
	}

	if cols <= 1 && rows <= 1 {
		return Size{Width: 1, Height: 1}, nil // nothing to merge
	}

	topLeft := cellRef.CellName()
	bottomRight := NewCellRef(cellRef.Sheet, cellRef.Row+rows-1, cellRef.Col+cols-1).CellName()

	if err := transformer.MergeCells(cellRef.Sheet, topLeft, bottomRight); err != nil {
		return ZeroSize, fmt.Errorf("merge cells %s:%s: %w", topLeft, bottomRight, err)
	}

	return Size{Width: cols, Height: rows}, nil
}

// toInt converts any numeric value to int.
func toInt(v any) int {
	switch n := v.(type) {
	case int:
		return n
	case int64:
		return int(n)
	case float64:
		return int(n)
	case float32:
		return int(n)
	case string:
		if i, err := strconv.Atoi(n); err == nil {
			return i
		}
	}
	return 1
}
