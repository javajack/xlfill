package xlfill

import (
	"fmt"
	"strings"
)

// Describe parses a template and returns a human-readable tree showing
// the area hierarchy, commands, and expressions found in cells.
// Useful for debugging templates during development.
func Describe(templatePath string, opts ...Option) (string, error) {
	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.Describe()
}

// Describe opens the template, parses its structure, and returns a
// human-readable tree of areas, commands, and expressions.
func (f *Filler) Describe() (string, error) {
	tx, err := f.openTemplate()
	if err != nil {
		return "", err
	}
	defer tx.Close()

	areas, err := f.BuildAreas(tx)
	if err != nil {
		return "", fmt.Errorf("build areas: %w", err)
	}

	var b strings.Builder
	b.WriteString("Template: ")
	if f.opts.templatePath != "" {
		b.WriteString(f.opts.templatePath)
	} else {
		b.WriteString("<reader>")
	}
	b.WriteByte('\n')

	for _, area := range areas {
		f.describeArea(&b, area, tx, 0)
	}
	return b.String(), nil
}

// describeArea recursively writes a tree description of an area and its commands.
func (f *Filler) describeArea(b *strings.Builder, area *Area, tx Transformer, indent int) {
	prefix := strings.Repeat("  ", indent)

	// Area header: Sheet1!A1:C10 area (3x10)
	lastCell := NewCellRef(
		area.StartCell.Sheet,
		area.StartCell.Row+area.AreaSize.Height-1,
		area.StartCell.Col+area.AreaSize.Width-1,
	)
	fmt.Fprintf(b, "%s%s:%s area %s\n", prefix, area.StartCell, lastCell.CellName(), area.AreaSize)

	// Collect child command cell ranges to skip when listing expressions
	childRanges := make([][4]int, 0, len(area.Bindings))
	for _, bind := range area.Bindings {
		childRanges = append(childRanges, [4]int{
			bind.StartRef.Row,
			bind.StartRef.Col,
			bind.StartRef.Row + bind.Size.Height - 1,
			bind.StartRef.Col + bind.Size.Width - 1,
		})
	}

	// Scan cells for expressions (skip cells covered by child commands)
	notationBegin := f.opts.notationBegin
	notationEnd := f.opts.notationEnd
	var exprs []string

	for row := 0; row < area.AreaSize.Height; row++ {
		for col := 0; col < area.AreaSize.Width; col++ {
			absRow := area.StartCell.Row + row
			absCol := area.StartCell.Col + col
			if inChildRange(absRow, absCol, childRanges) {
				continue
			}
			ref := NewCellRef(area.StartCell.Sheet, absRow, absCol)
			cd := tx.GetCellData(ref)
			if cd == nil {
				continue
			}
			if strVal, ok := cd.Value.(string); ok && strings.Contains(strVal, notationBegin) {
				exprs = append(exprs, fmt.Sprintf("%s%s: %s", prefix+"    ", ref.CellName(), strVal))
			}
			if cd.Formula != "" && strings.Contains(cd.Formula, notationBegin) {
				exprs = append(exprs, fmt.Sprintf("%s%s: =%s", prefix+"    ", ref.CellName(), cd.Formula))
			}
		}
	}
	if len(exprs) > 0 {
		fmt.Fprintf(b, "%s  Expressions:\n", prefix)
		for _, e := range exprs {
			b.WriteString(e)
			b.WriteByte('\n')
		}
	}

	// Commands
	if len(area.Bindings) > 0 {
		fmt.Fprintf(b, "%s  Commands:\n", prefix)
		for _, bind := range area.Bindings {
			attrs := describeCommandAttrs(bind.Command)
			fmt.Fprintf(b, "%s    %s %s %s%s\n", prefix, bind.StartRef, bind.Command.Name(), bind.Size, attrs)

			// Recurse into child area
			if childArea := getCommandArea(bind.Command); childArea != nil {
				f.describeArea(b, childArea, tx, indent+3)
			}
		}
	}
	_ = notationEnd
}

// inChildRange checks if a cell (row, col) falls within any child command range.
func inChildRange(row, col int, ranges [][4]int) bool {
	for _, r := range ranges {
		if row >= r[0] && row <= r[2] && col >= r[1] && col <= r[3] {
			return true
		}
	}
	return false
}

// describeCommandAttrs returns a string of key command attributes for display.
func describeCommandAttrs(cmd Command) string {
	var parts []string
	switch c := cmd.(type) {
	case *EachCommand:
		parts = append(parts, fmt.Sprintf("items=%q", c.Items))
		parts = append(parts, fmt.Sprintf("var=%q", c.Var))
		if c.VarIndex != "" {
			parts = append(parts, fmt.Sprintf("varIndex=%q", c.VarIndex))
		}
		if c.Direction != "" && c.Direction != "DOWN" {
			parts = append(parts, fmt.Sprintf("direction=%q", c.Direction))
		}
		if c.Select != "" {
			parts = append(parts, fmt.Sprintf("select=%q", c.Select))
		}
		if c.OrderBy != "" {
			parts = append(parts, fmt.Sprintf("orderBy=%q", c.OrderBy))
		}
		if c.GroupBy != "" {
			parts = append(parts, fmt.Sprintf("groupBy=%q", c.GroupBy))
		}
		if c.MultiSheet != "" {
			parts = append(parts, fmt.Sprintf("multiSheet=%q", c.MultiSheet))
		}
	case *IfCommand:
		parts = append(parts, fmt.Sprintf("condition=%q", c.Condition))
	case *GridCommand:
		parts = append(parts, fmt.Sprintf("headers=%q", c.Headers))
		parts = append(parts, fmt.Sprintf("data=%q", c.Data))
		if c.Props != "" {
			parts = append(parts, fmt.Sprintf("props=%q", c.Props))
		}
	case *ImageCommand:
		parts = append(parts, fmt.Sprintf("src=%q", c.Src))
		if c.ImageType != "" {
			parts = append(parts, fmt.Sprintf("imageType=%q", c.ImageType))
		}
	case *MergeCellsCommand:
		if c.Cols != "" {
			parts = append(parts, fmt.Sprintf("cols=%q", c.Cols))
		}
		if c.Rows != "" {
			parts = append(parts, fmt.Sprintf("rows=%q", c.Rows))
		}
	case *UpdateCellCommand:
		parts = append(parts, fmt.Sprintf("updater=%q", c.Updater))
	case *AutoRowHeightCommand:
		// no extra attributes
	case *RepeatCommand:
		parts = append(parts, fmt.Sprintf("count=%q", c.Count))
		if c.Var != "" {
			parts = append(parts, fmt.Sprintf("var=%q", c.Var))
		}
		if c.Direction != "" && c.Direction != "DOWN" {
			parts = append(parts, fmt.Sprintf("direction=%q", c.Direction))
		}
	case *DataValidationCommand:
		parts = append(parts, fmt.Sprintf("type=%q", c.ValidationType))
		if c.Source != "" {
			parts = append(parts, fmt.Sprintf("source=%q", c.Source))
		}
	case *TableCommand:
		parts = append(parts, fmt.Sprintf("name=%q", c.TableName))
		parts = append(parts, fmt.Sprintf("style=%q", c.Style))
	case *ConditionalFormatCommand:
		parts = append(parts, fmt.Sprintf("type=%q", c.FormatType))
	case *GroupCommand:
		if c.Collapsed != "false" {
			parts = append(parts, fmt.Sprintf("collapsed=%q", c.Collapsed))
		}
	case *ChartCommand:
		parts = append(parts, fmt.Sprintf("type=%q", c.ChartType))
		if c.Title != "" {
			parts = append(parts, fmt.Sprintf("title=%q", c.Title))
		}
	case *DefinedNameCommand:
		parts = append(parts, fmt.Sprintf("name=%q", c.DefinedNameValue))
		if c.Scope != "workbook" {
			parts = append(parts, fmt.Sprintf("scope=%q", c.Scope))
		}
	case *SparklineCommand:
		parts = append(parts, fmt.Sprintf("type=%q", c.SparkType))
		parts = append(parts, fmt.Sprintf("data=%q", c.DataRange))
	case *PageBreakCommand:
		// no extra attributes
	case *AutoColWidthCommand:
		// no extra attributes
	case *FreezePanesCommand:
		if c.FreezeRow != "" {
			parts = append(parts, fmt.Sprintf("row=%q", c.FreezeRow))
		}
		if c.FreezeCol != "" {
			parts = append(parts, fmt.Sprintf("col=%q", c.FreezeCol))
		}
	case *ProtectCommand:
		if c.Password != "" {
			parts = append(parts, "password=***")
		}
		if c.AllowSort == "true" {
			parts = append(parts, "allowSort=\"true\"")
		}
		if c.AllowFilter == "true" {
			parts = append(parts, "allowFilter=\"true\"")
		}
	case *IncludeCommand:
		parts = append(parts, fmt.Sprintf("template=%q", c.TemplatePath))
		if c.SourceSheet != "" {
			parts = append(parts, fmt.Sprintf("sheet=%q", c.SourceSheet))
		}
		parts = append(parts, fmt.Sprintf("area=%q", c.SourceArea))
	}
	if len(parts) == 0 {
		return ""
	}
	return " " + strings.Join(parts, " ")
}
