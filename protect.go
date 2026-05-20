package xlfill

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// ProtectCommand implements the jx:protect command.
// It processes its inner area, then protects the sheet with the specified options.
type ProtectCommand struct {
	Password                 string
	AllowSort                string
	AllowFilter              string
	AllowSelectUnlockedCells string
	AllowSelectLockedCells   string
	Area                     *Area
}

func (c *ProtectCommand) Name() string    { return "protect" }
func (c *ProtectCommand) Reset()          {}
func (c *ProtectCommand) GetArea() *Area  { return c.Area }
func (c *ProtectCommand) SetArea(a *Area) { c.Area = a }

// newProtectCommandFromAttrs creates a ProtectCommand from parsed attributes.
func newProtectCommandFromAttrs(attrs map[string]string) (Command, error) {
	return &ProtectCommand{
		Password:                 attrs["password"],
		AllowSort:                attrs["allowSort"],
		AllowFilter:              attrs["allowFilter"],
		AllowSelectUnlockedCells: attrs["allowSelectUnlockedCells"],
		AllowSelectLockedCells:   attrs["allowSelectLockedCells"],
	}, nil
}

// ApplyAt processes the inner area, then protects the sheet.
func (c *ProtectCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
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

	opts := &excelize.SheetProtectionOptions{
		Password:            c.Password,
		SelectLockedCells:   c.AllowSelectLockedCells == "true",
		SelectUnlockedCells: c.AllowSelectUnlockedCells == "true",
		Sort:                c.AllowSort == "true",
		AutoFilter:          c.AllowFilter == "true",
	}
	if err := etx.file.ProtectSheet(cellRef.Sheet, opts); err != nil {
		return size, fmt.Errorf("protect sheet: %w", err)
	}

	return size, nil
}
