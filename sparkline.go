package xlfill

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// SparklineCommand implements the jx:sparkline command.
// It processes its inner area normally, then registers a deferred action
// that adds an Excel sparkline at the anchor cell.
type SparklineCommand struct {
	SparkType string // line, column, winLoss
	DataRange string // data range for sparkline (e.g., "Sheet1!A1:A10")
	Color     string // sparkline series color
	Area      *Area
}

func (c *SparklineCommand) Name() string { return "sparkline" }
func (c *SparklineCommand) Reset()       {}

// newSparklineCommandFromAttrs creates a SparklineCommand from parsed attributes.
func newSparklineCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &SparklineCommand{
		SparkType: attrs["type"],
		DataRange: attrs["data"],
		Color:     attrs["color"],
	}
	if cmd.SparkType == "" {
		cmd.SparkType = "line"
	}
	if cmd.DataRange == "" {
		return nil, fmt.Errorf("sparkline command requires 'data' attribute")
	}
	return cmd, nil
}

// sparkTypeMap maps user-friendly sparkline type names to excelize type strings.
var sparkTypeMap = map[string]string{
	"line":    "line",
	"column":  "column",
	"winLoss": "win_loss",
}

// ApplyAt processes the inner area, then registers a deferred action for sparkline creation.
func (c *SparklineCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
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
	anchorCell := cellRef.CellName()

	sparkType := c.SparkType
	dataRange := c.DataRange
	color := c.Color

	// Register deferred action
	if err := ctx.RegisterDeferred(DeferredAction{
		Name:     "sparkline",
		Sheet:    sheet,
		StartRow: cellRef.Row,
		StartCol: cellRef.Col,
		EndRow:   cellRef.Row + size.Height - 1,
		EndCol:   cellRef.Col + size.Width - 1,
		Execute: func(tx *ExcelizeTransformer) error {
			st, ok := sparkTypeMap[sparkType]
			if !ok {
				st = "line" // default
			}

			opts := &excelize.SparklineOptions{
				Location: []string{anchorCell},
				Range:    []string{dataRange},
				Type:     st,
				Style:    1,
			}

			if color != "" {
				opts.SeriesColor = color
			}

			return tx.file.AddSparkline(sheet, opts)
		},
	}); err != nil {
		return ZeroSize, fmt.Errorf("sparkline: %w", err)
	}

	return size, nil
}
