package xlfill

import (
	"fmt"
	"strings"
)

// CellRef represents a single cell reference in an Excel workbook.
type CellRef struct {
	Sheet string // sheet name (empty = current sheet)
	Row   int    // 0-based row index
	Col   int    // 0-based column index
}

// NewCellRef creates a CellRef with explicit sheet, row, col.
func NewCellRef(sheet string, row, col int) CellRef {
	return CellRef{Sheet: sheet, Row: row, Col: col}
}

// ParseCellRef parses a cell reference string like "A1", "Sheet1!B5", or "$A$1".
func ParseCellRef(s string) (CellRef, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return CellRef{}, fmt.Errorf("empty cell reference")
	}

	var sheet string
	cellPart := s

	if idx := strings.LastIndex(s, "!"); idx >= 0 {
		sheet = strings.Trim(s[:idx], "'")
		cellPart = s[idx+1:]
	}

	cellPart = strings.ReplaceAll(cellPart, "$", "")
	if cellPart == "" {
		return CellRef{}, fmt.Errorf("invalid cell reference: %q", s)
	}

	col, row, err := parseCellName(cellPart)
	if err != nil {
		return CellRef{}, fmt.Errorf("invalid cell reference %q: %w", s, err)
	}

	return CellRef{Sheet: sheet, Row: row, Col: col}, nil
}

// parseCellName parses "A1" into col=0, row=0.
func parseCellName(name string) (col, row int, err error) {
	if len(name) == 0 {
		return 0, 0, fmt.Errorf("empty cell name")
	}

	i := 0
	for i < len(name) && isAlpha(name[i]) {
		i++
	}
	if i == 0 || i == len(name) {
		return 0, 0, fmt.Errorf("invalid cell name: %q", name)
	}

	colStr := name[:i]
	rowStr := name[i:]

	col, err = NameToCol(colStr)
	if err != nil {
		return 0, 0, err
	}

	rowNum := 0
	for _, ch := range rowStr {
		if ch < '0' || ch > '9' {
			return 0, 0, fmt.Errorf("invalid row in cell name: %q", name)
		}
		rowNum = rowNum*10 + int(ch-'0')
	}
	if rowNum < 1 {
		return 0, 0, fmt.Errorf("invalid row number in cell name: %q", name)
	}

	return col, rowNum - 1, nil // convert 1-based row to 0-based
}

func isAlpha(b byte) bool {
	return (b >= 'A' && b <= 'Z') || (b >= 'a' && b <= 'z')
}

// String formats the CellRef as "Sheet1!A1" or "A1" if no sheet.
func (c CellRef) String() string {
	name := c.CellName()
	if c.Sheet != "" {
		return c.Sheet + "!" + name
	}
	return name
}

// CellName returns just the cell part like "A1" without sheet name.
func (c CellRef) CellName() string {
	return ColToName(c.Col) + fmt.Sprintf("%d", c.Row+1)
}

// ColToName converts a 0-based column index to a column name.
// 0→"A", 25→"Z", 26→"AA", 702→"AAA"
func ColToName(col int) string {
	result := ""
	col++ // convert to 1-based for algorithm
	for col > 0 {
		col-- // adjust for 0-indexed letter
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// NameToCol converts a column name to a 0-based column index.
// "A"→0, "Z"→25, "AA"→26
func NameToCol(name string) (int, error) {
	name = strings.ToUpper(name)
	if name == "" {
		return 0, fmt.Errorf("empty column name")
	}
	col := 0
	for _, ch := range name {
		if ch < 'A' || ch > 'Z' {
			return 0, fmt.Errorf("invalid column name: %q", name)
		}
		col = col*26 + int(ch-'A') + 1
	}
	return col - 1, nil
}

// AreaRef represents a rectangular area defined by two cell references.
type AreaRef struct {
	First CellRef
	Last  CellRef
}

// NewAreaRef creates an AreaRef from two cell references.
func NewAreaRef(first, last CellRef) AreaRef {
	return AreaRef{First: first, Last: last}
}

// ParseAreaRef parses an area reference string like "A1:C5" or "Sheet1!A1:C5".
func ParseAreaRef(s string) (AreaRef, error) {
	s = strings.TrimSpace(s)
	parts := strings.SplitN(s, ":", 2)
	if len(parts) != 2 {
		return AreaRef{}, fmt.Errorf("invalid area reference (missing ':'): %q", s)
	}

	first, err := ParseCellRef(parts[0])
	if err != nil {
		return AreaRef{}, fmt.Errorf("invalid area reference %q: %w", s, err)
	}

	last, err := ParseCellRef(parts[1])
	if err != nil {
		return AreaRef{}, fmt.Errorf("invalid area reference %q: %w", s, err)
	}

	// Inherit sheet name from first cell if last doesn't have one
	if last.Sheet == "" && first.Sheet != "" {
		last.Sheet = first.Sheet
	}

	return AreaRef{First: first, Last: last}, nil
}

// String formats the AreaRef as "Sheet1!A1:C5" or "A1:C5".
func (a AreaRef) String() string {
	if a.First.Sheet != "" && a.First.Sheet == a.Last.Sheet {
		return a.First.Sheet + "!" + a.First.CellName() + ":" + a.Last.CellName()
	}
	return a.First.String() + ":" + a.Last.String()
}

// Size returns the dimensions of the area.
func (a AreaRef) Size() Size {
	return Size{
		Width:  a.Last.Col - a.First.Col + 1,
		Height: a.Last.Row - a.First.Row + 1,
	}
}

// Contains returns true if the given cell reference is within this area.
func (a AreaRef) Contains(ref CellRef) bool {
	if a.First.Sheet != "" && a.First.Sheet != ref.Sheet {
		return false
	}
	return ref.Row >= a.First.Row && ref.Row <= a.Last.Row &&
		ref.Col >= a.First.Col && ref.Col <= a.Last.Col
}

// SheetName returns the sheet name of this area (from First cell).
func (a AreaRef) SheetName() string {
	return a.First.Sheet
}

// Size represents width (columns) and height (rows).
type Size struct {
	Width  int
	Height int
}

// ZeroSize is a Size with zero width and height.
var ZeroSize = Size{Width: 0, Height: 0}

// String formats the Size as "(WxH)".
func (s Size) String() string {
	return fmt.Sprintf("(%dx%d)", s.Width, s.Height)
}

// Add returns a new Size with both dimensions added.
func (s Size) Add(other Size) Size {
	return Size{Width: s.Width + other.Width, Height: s.Height + other.Height}
}

// Minus returns a new Size with both dimensions subtracted.
func (s Size) Minus(other Size) Size {
	return Size{Width: s.Width - other.Width, Height: s.Height - other.Height}
}

// SafeSheetName sanitizes a string for use as an Excel sheet name.
// It replaces forbidden characters ([]*?/\:) with underscore and truncates to 31 chars.
func SafeSheetName(name string) string {
	forbidden := []rune{'/', '\\', ':', '*', '?', '[', ']'}
	runes := []rune(name)
	for i, r := range runes {
		for _, f := range forbidden {
			if r == f {
				runes[i] = '_'
				break
			}
		}
	}
	if len(runes) > 31 {
		runes = runes[:31]
	}
	return string(runes)
}
