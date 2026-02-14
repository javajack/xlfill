package xlfill

import "io"

// Transformer abstracts Excel I/O operations. It reads template data into memory
// and provides methods to transform cells from source to target positions.
type Transformer interface {
	// Cell data access
	GetCellData(ref CellRef) *CellData
	GetCommentedCells() []*CellData
	GetFormulaCells() []*CellData

	// Cell transformation
	Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error
	ClearCell(ref CellRef) error
	SetFormula(ref CellRef, formula string) error
	SetCellValue(ref CellRef, value any) error

	// Target tracking for formula processing
	GetTargetCellRef(src CellRef) []CellRef
	ResetTargetCellRefs()

	// Sheet data
	GetSheetNames() []string
	GetColumnWidth(sheet string, col int) float64
	GetRowHeight(sheet string, row int) float64

	// Sheet operations
	DeleteSheet(name string) error
	SetHidden(name string, hidden bool) error
	CopySheet(src, dst string) error

	// Image/merge/hyperlink
	AddImage(sheet string, cell string, imgBytes []byte, imgType string, scaleX, scaleY float64) error
	MergeCells(sheet, topLeft, bottomRight string) error
	SetCellHyperLink(ref CellRef, url, display string) error

	// Workbook properties
	SetRecalculateOnOpen(recalc bool) error

	// I/O
	Write(w io.Writer) error
	Close() error
}

// SheetData holds in-memory data for a single sheet.
type SheetData struct {
	Name         string
	ColumnWidths map[int]float64
	Rows         map[int]*RowData
}

// RowData holds in-memory data for a single row.
type RowData struct {
	Height float64
	Cells  map[int]*CellData
}
