package xlfill

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// DataValidationCommand implements the jx:dataValidation command.
// It processes its inner area normally, then registers a deferred action
// that applies Excel data validation on the expanded output range.
type DataValidationCommand struct {
	ValidationType string // list, integer, decimal, date, custom
	Source         string // comma-separated values for list, or formula for custom
	Min            string // minimum value (for integer/decimal/date)
	Max            string // maximum value (for integer/decimal/date)
	AllowBlank     string // "true" (default) or "false"
	ShowError      string // "true" or "false"
	ErrorTitle     string // error dialog title
	ErrorMsg       string // error dialog message
	Area           *Area
}

func (c *DataValidationCommand) Name() string { return "dataValidation" }
func (c *DataValidationCommand) Reset()       {}

// newDataValidationCommandFromAttrs creates a DataValidationCommand from parsed attributes.
func newDataValidationCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &DataValidationCommand{
		ValidationType: attrs["type"],
		Source:         attrs["source"],
		Min:            attrs["min"],
		Max:            attrs["max"],
		AllowBlank:     attrs["allowBlank"],
		ShowError:      attrs["showError"],
		ErrorTitle:     attrs["errorTitle"],
		ErrorMsg:       attrs["error"],
	}
	if cmd.ValidationType == "" {
		return nil, fmt.Errorf("dataValidation command requires 'type' attribute")
	}
	if cmd.AllowBlank == "" {
		cmd.AllowBlank = "true"
	}
	return cmd, nil
}

// ApplyAt processes the inner area, then registers a deferred action for data validation.
func (c *DataValidationCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
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

	validationType := c.ValidationType
	source := c.Source
	minVal := c.Min
	maxVal := c.Max
	allowBlank := c.AllowBlank != "false"
	showError := c.ShowError == "true"
	errorTitle := c.ErrorTitle
	errorMsg := c.ErrorMsg

	// Register deferred action
	ctx.RegisterDeferred(DeferredAction{
		Name:     "dataValidation",
		Sheet:    sheet,
		StartRow: startRow,
		StartCol: startCol,
		EndRow:   endRow,
		EndCol:   endCol,
		Execute: func(tx *ExcelizeTransformer) error {
			dv := excelize.NewDataValidation(allowBlank)

			startCellName := NewCellRef(sheet, startRow, startCol).CellName()
			endCellName := NewCellRef(sheet, endRow, endCol).CellName()
			dv.Sqref = fmt.Sprintf("%s:%s", startCellName, endCellName)

			switch validationType {
			case "list":
				if err := dv.SetDropList(strings.Split(source, ",")); err != nil {
					return fmt.Errorf("set drop list: %w", err)
				}
			case "integer":
				if err := dv.SetRange(minVal, maxVal, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween); err != nil {
					return fmt.Errorf("set integer range: %w", err)
				}
			case "decimal":
				if err := dv.SetRange(minVal, maxVal, excelize.DataValidationTypeDecimal, excelize.DataValidationOperatorBetween); err != nil {
					return fmt.Errorf("set decimal range: %w", err)
				}
			case "date":
				if err := dv.SetRange(minVal, maxVal, excelize.DataValidationTypeDate, excelize.DataValidationOperatorBetween); err != nil {
					return fmt.Errorf("set date range: %w", err)
				}
			case "custom":
				dv.Type = "custom"
				dv.Formula1 = source
			default:
				return fmt.Errorf("unsupported data validation type: %q", validationType)
			}

			if showError {
				dv.SetError(excelize.DataValidationErrorStyleStop, errorTitle, errorMsg)
			}

			return tx.file.AddDataValidation(sheet, dv)
		},
	})

	return size, nil
}
