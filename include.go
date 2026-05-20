package xlfill

import (
	"fmt"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// MaxIncludeDepth limits the nesting depth of jx:include commands to prevent infinite recursion.
const MaxIncludeDepth = 10

// validateIncludePath rejects absolute paths and parent directory traversal.
func validateIncludePath(path string) error {
	if filepath.IsAbs(path) {
		return fmt.Errorf("absolute paths not allowed: %q", path)
	}
	cleaned := filepath.Clean(path)
	if strings.HasPrefix(cleaned, "..") {
		return fmt.Errorf("path traversal not allowed: %q", path)
	}
	return nil
}

// IncludeCommand implements the jx:include command.
// It opens an external template file, reads cells from a specified range,
// and copies them to the target position.
type IncludeCommand struct {
	TemplatePath string // file path to the included template
	SourceSheet  string // source sheet name (defaults to first sheet)
	SourceArea   string // source range e.g. "A1:F3"
	Area         *Area
}

func (c *IncludeCommand) Name() string    { return "include" }
func (c *IncludeCommand) Reset()          {}
func (c *IncludeCommand) GetArea() *Area  { return c.Area }
func (c *IncludeCommand) SetArea(a *Area) { c.Area = a }

// newIncludeCommandFromAttrs creates an IncludeCommand from parsed attributes.
func newIncludeCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &IncludeCommand{
		TemplatePath: attrs["template"],
		SourceSheet:  attrs["sheet"],
		SourceArea:   attrs["area"],
	}
	if cmd.TemplatePath == "" {
		return nil, fmt.Errorf("include: template attribute is required")
	}
	if cmd.SourceArea == "" {
		return nil, fmt.Errorf("include: area attribute is required")
	}
	return cmd, nil
}

// ApplyAt opens the included template, reads cells from the specified range,
// and copies them to the target position.
func (c *IncludeCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	// Check include depth to prevent infinite recursion
	if ctx.includeDepth >= MaxIncludeDepth {
		return ZeroSize, fmt.Errorf("include: maximum nesting depth (%d) exceeded", MaxIncludeDepth)
	}
	ctx.includeDepth++
	defer func() { ctx.includeDepth-- }()

	// Evaluate template path (could be an expression)
	tmplPath := c.TemplatePath
	if val, err := ctx.Evaluate(tmplPath); err == nil {
		if s, ok := val.(string); ok {
			tmplPath = s
		}
	}

	// Validate path to prevent traversal attacks
	if err := validateIncludePath(tmplPath); err != nil {
		return ZeroSize, fmt.Errorf("include: %w", err)
	}

	// Open the included template
	incFile, err := excelize.OpenFile(tmplPath)
	if err != nil {
		return ZeroSize, fmt.Errorf("include: open %q: %w", tmplPath, err)
	}
	defer incFile.Close()

	srcSheet := c.SourceSheet
	if srcSheet == "" {
		sheets := incFile.GetSheetList()
		if len(sheets) > 0 {
			srcSheet = sheets[0]
		}
	}

	// Parse the source area
	srcArea, err := ParseAreaRef(c.SourceArea)
	if err != nil {
		return ZeroSize, fmt.Errorf("include: parse area %q: %w", c.SourceArea, err)
	}

	etx := unwrapExcelizeTransformer(transformer)
	if etx == nil {
		return ZeroSize, nil
	}

	// Copy cells from included file to target position
	areaH := srcArea.Last.Row - srcArea.First.Row + 1
	areaW := srcArea.Last.Col - srcArea.First.Col + 1
	for row := 0; row < areaH; row++ {
		for col := 0; col < areaW; col++ {
			srcCellName := ColToName(srcArea.First.Col+col) + strconv.Itoa(srcArea.First.Row+row+1)
			val, err := incFile.GetCellValue(srcSheet, srcCellName)
			if err != nil {
				return ZeroSize, fmt.Errorf("include: read cell %s: %w", srcCellName, err)
			}
			targetRef := NewCellRef(cellRef.Sheet, cellRef.Row+row, cellRef.Col+col)
			if err := etx.file.SetCellValue(cellRef.Sheet, targetRef.CellName(), val); err != nil {
				return ZeroSize, fmt.Errorf("include: write cell %s: %w", targetRef.CellName(), err)
			}
		}
	}

	return Size{Width: areaW, Height: areaH}, nil
}
