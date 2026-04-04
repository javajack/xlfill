package xlfill

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// ConditionalFormatCommand implements the jx:conditionalFormat command.
// It processes its inner area normally, then registers a deferred action
// that applies conditional formatting on the expanded output range.
type ConditionalFormatCommand struct {
	FormatType string // colorScale, dataBar, iconSet, cellIs, top10, aboveAverage
	Operator   string // for cellIs: >, <, ==, between, etc.
	Value      string // comparison value
	Format     string // style format (not used for dataBar/colorScale)
	MinColor   string // for colorScale
	MaxColor   string // for colorScale
	Color      string // for dataBar
	IconStyle  string // for iconSet (e.g., "3Arrows")
	Area       *Area
}

func (c *ConditionalFormatCommand) Name() string { return "conditionalFormat" }
func (c *ConditionalFormatCommand) Reset()       {}

// newConditionalFormatCommandFromAttrs creates a ConditionalFormatCommand from parsed attributes.
func newConditionalFormatCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &ConditionalFormatCommand{
		FormatType: attrs["type"],
		Operator:   attrs["operator"],
		Value:      attrs["value"],
		Format:     attrs["format"],
		MinColor:   attrs["minColor"],
		MaxColor:   attrs["maxColor"],
		Color:      attrs["color"],
		IconStyle:  attrs["iconStyle"],
	}
	if cmd.FormatType == "" {
		return nil, fmt.Errorf("conditionalFormat command requires 'type' attribute")
	}
	return cmd, nil
}

// ApplyAt processes the inner area, then registers a deferred action for conditional formatting.
func (c *ConditionalFormatCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
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

	cfType := c.FormatType
	operator := c.Operator
	value := c.Value
	minColor := c.MinColor
	maxColor := c.MaxColor
	color := c.Color
	iconStyle := c.IconStyle

	// Register deferred action
	if err := ctx.RegisterDeferred(DeferredAction{
		Name:     "conditionalFormat",
		Sheet:    sheet,
		StartRow: startRow,
		StartCol: startCol,
		EndRow:   endRow,
		EndCol:   endCol,
		Execute: func(tx *ExcelizeTransformer) error {
			startCellName := NewCellRef(sheet, startRow, startCol).CellName()
			endCellName := NewCellRef(sheet, endRow, endCol).CellName()
			rangeStr := fmt.Sprintf("%s:%s", startCellName, endCellName)

			var opts []excelize.ConditionalFormatOptions

			switch cfType {
			case "dataBar":
				barColor := color
				if barColor == "" {
					barColor = "#638EC6"
				}
				opts = append(opts, excelize.ConditionalFormatOptions{
					Type:     "data_bar",
					Criteria: "=",
					MinType:  "min",
					MaxType:  "max",
					BarColor: barColor,
				})
			case "colorScale":
				mc := minColor
				if mc == "" {
					mc = "#F8696B"
				}
				xc := maxColor
				if xc == "" {
					xc = "#63BE7B"
				}
				opts = append(opts, excelize.ConditionalFormatOptions{
					Type:     "2_color_scale",
					Criteria: "=",
					MinType:  "min",
					MaxType:  "max",
					MinColor: mc,
					MaxColor: xc,
				})
			case "iconSet":
				is := iconStyle
				if is == "" {
					is = "3Arrows"
				}
				opts = append(opts, excelize.ConditionalFormatOptions{
					Type:      "icon_set",
					IconStyle: is,
				})
			case "cellIs":
				criteria := mapOperatorToCriteria(operator)
				opts = append(opts, excelize.ConditionalFormatOptions{
					Type:     "cell",
					Criteria: criteria,
					Value:    value,
				})
			case "top10":
				opts = append(opts, excelize.ConditionalFormatOptions{
					Type:     "top",
					Criteria: "=",
					Value:    value,
				})
			case "aboveAverage":
				opts = append(opts, excelize.ConditionalFormatOptions{
					Type:     "average",
					Criteria: "=",
				})
			default:
				return fmt.Errorf("unsupported conditional format type: %q", cfType)
			}

			return tx.file.SetConditionalFormat(sheet, rangeStr, opts)
		},
	}); err != nil {
		return ZeroSize, fmt.Errorf("conditionalFormat: %w", err)
	}

	return size, nil
}

// mapOperatorToCriteria maps user-friendly operator strings to excelize criteria strings.
func mapOperatorToCriteria(op string) string {
	switch op {
	case ">", "greaterThan":
		return ">"
	case "<", "lessThan":
		return "<"
	case ">=", "greaterThanOrEqual":
		return ">="
	case "<=", "lessThanOrEqual":
		return "<="
	case "==", "equal":
		return "=="
	case "!=", "notEqual":
		return "!="
	case "between":
		return "between"
	default:
		return op
	}
}
