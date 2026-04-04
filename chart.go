package xlfill

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// ChartCommand implements the jx:chart command.
// It processes its inner area normally, then registers a deferred action
// that adds an Excel chart anchored at the start cell of the area.
type ChartCommand struct {
	ChartType  string // bar, col, line, pie, scatter, area, doughnut, radar
	Title      string // chart title
	Series     string // data series range (e.g., "Sheet1!$B$2:$B$10")
	Categories string // categories range (e.g., "Sheet1!$A$2:$A$10")
	Width      string // chart width in pixels (default 480)
	Height     string // chart height in pixels (default 290)
	Area       *Area
}

func (c *ChartCommand) Name() string { return "chart" }
func (c *ChartCommand) Reset()       {}

// newChartCommandFromAttrs creates a ChartCommand from parsed attributes.
func newChartCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &ChartCommand{
		ChartType:  attrs["type"],
		Title:      attrs["title"],
		Series:     attrs["series"],
		Categories: attrs["categories"],
		Width:      attrs["width"],
		Height:     attrs["height"],
	}
	if cmd.ChartType == "" {
		return nil, fmt.Errorf("chart command requires 'type' attribute")
	}
	if cmd.Width == "" {
		cmd.Width = "480"
	}
	if cmd.Height == "" {
		cmd.Height = "290"
	}
	return cmd, nil
}

// chartTypeMap maps user-friendly chart type names to excelize ChartType constants.
var chartTypeMap = map[string]excelize.ChartType{
	"bar":      excelize.Bar,
	"col":      excelize.Col,
	"line":     excelize.Line,
	"pie":      excelize.Pie,
	"scatter":  excelize.Scatter,
	"area":     excelize.Area,
	"doughnut": excelize.Doughnut,
	"radar":    excelize.Radar,
}

// ApplyAt processes the inner area, then registers a deferred action for chart creation.
func (c *ChartCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
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

	chartTypeName := c.ChartType
	title := c.Title
	series := c.Series
	categories := c.Categories

	width, _ := strconv.ParseUint(c.Width, 10, 32)
	if width == 0 {
		width = 480
	}
	height, _ := strconv.ParseUint(c.Height, 10, 32)
	if height == 0 {
		height = 290
	}

	// Validate chart type at creation time (fail fast)
	if _, ok := chartTypeMap[chartTypeName]; !ok {
		return ZeroSize, fmt.Errorf("unsupported chart type: %q", chartTypeName)
	}

	// Register deferred action
	if err := ctx.RegisterDeferred(DeferredAction{
		Name:     "chart",
		Sheet:    sheet,
		StartRow: cellRef.Row,
		StartCol: cellRef.Col,
		EndRow:   cellRef.Row + size.Height - 1,
		EndCol:   cellRef.Col + size.Width - 1,
		Execute: func(tx *ExcelizeTransformer) error {
			ct := chartTypeMap[chartTypeName]

			chartSeries := []excelize.ChartSeries{
				{
					Name:       title,
					Categories: categories,
					Values:     series,
				},
			}

			return tx.file.AddChart(sheet, anchorCell, &excelize.Chart{
				Type:   ct,
				Series: chartSeries,
				Title:  []excelize.RichTextRun{{Text: title}},
				Dimension: excelize.ChartDimension{
					Width:  uint(width),
					Height: uint(height),
				},
			})
		},
	}); err != nil {
		return ZeroSize, fmt.Errorf("chart: %w", err)
	}

	return size, nil
}
