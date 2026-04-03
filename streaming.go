package xlfill

import (
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

// streamCell holds a single buffered cell for streaming output.
type streamCell struct {
	value   any
	styleID int
	formula string
}

// StreamingTransformer writes output rows via excelize StreamWriter instead of
// holding the entire output in memory. Template data is read from an embedded
// ExcelizeTransformer (read-only after init). Output cells are buffered per row
// and flushed to the StreamWriter in ascending row order.
//
// Limitations of streaming mode:
//   - Formula post-processing (reference remapping) is not supported
//   - Hyperlinks are not supported (silently ignored)
//   - Images are not supported (returns error)
//   - Rows must be written in ascending order (guaranteed by area processing)
type StreamingTransformer struct {
	reader       *ExcelizeTransformer
	sw           *excelize.StreamWriter
	sheet        string
	mu           sync.Mutex
	rowBuf       map[int]map[int]*streamCell // row -> col -> cell
	nextFlushRow int                         // next row index to flush
	closed       bool
}

// NewStreamingTransformer creates a streaming transformer for the given sheet.
// It initializes the StreamWriter and copies column widths from the template.
func NewStreamingTransformer(reader *ExcelizeTransformer, sheet string) (*StreamingTransformer, error) {
	sw, err := reader.file.NewStreamWriter(sheet)
	if err != nil {
		return nil, fmt.Errorf("create stream writer for sheet %q: %w", sheet, err)
	}

	st := &StreamingTransformer{
		reader: reader,
		sw:     sw,
		sheet:  sheet,
		rowBuf: make(map[int]map[int]*streamCell),
	}

	// Copy column widths from template
	if sd, ok := reader.sheets[sheet]; ok {
		for col, w := range sd.ColumnWidths {
			colName := ColToName(col)
			colNum := col + 1
			sw.SetColWidth(colNum, colNum, w)
			_ = colName
		}
	}

	return st, nil
}

// bufferCell stores a cell value for later flushing.
func (st *StreamingTransformer) bufferCell(ref CellRef, value any, styleID int, formula string) error {
	st.mu.Lock()
	defer st.mu.Unlock()

	row := ref.Row

	// Flush all completed rows below this one
	for r := st.nextFlushRow; r < row; r++ {
		if err := st.flushRowLocked(r); err != nil {
			return err
		}
	}

	// Buffer this cell
	if st.rowBuf[row] == nil {
		st.rowBuf[row] = make(map[int]*streamCell)
	}
	st.rowBuf[row][ref.Col] = &streamCell{value: value, styleID: styleID, formula: formula}
	return nil
}

// flushRowLocked writes a buffered row to the StreamWriter. Caller must hold mu.
func (st *StreamingTransformer) flushRowLocked(row int) error {
	cells := st.rowBuf[row]
	if len(cells) == 0 {
		st.nextFlushRow = row + 1
		return nil
	}

	// Find min and max column to avoid padding unused leading columns
	minCol, maxCol := int(^uint(0)>>1), 0
	for col := range cells {
		if col < minCol {
			minCol = col
		}
		if col > maxCol {
			maxCol = col
		}
	}
	// Build values slice starting from minCol
	values := make([]interface{}, maxCol-minCol+1)
	for col := minCol; col <= maxCol; col++ {
		idx := col - minCol
		if cell, ok := cells[col]; ok {
			if cell.formula != "" {
				values[idx] = excelize.Cell{StyleID: cell.styleID, Formula: cell.formula}
			} else if cell.value != nil {
				values[idx] = excelize.Cell{StyleID: cell.styleID, Value: cell.value}
			} else {
				values[idx] = excelize.Cell{StyleID: cell.styleID}
			}
		}
	}

	startCell := ColToName(minCol) + strconv.Itoa(row+1)
	if err := st.sw.SetRow(startCell, values); err != nil {
		return fmt.Errorf("stream write row %d: %w", row+1, err)
	}

	delete(st.rowBuf, row)
	st.nextFlushRow = row + 1
	return nil
}

// Flush writes all remaining buffered rows and finalizes the StreamWriter.
func (st *StreamingTransformer) Flush() error {
	st.mu.Lock()
	defer st.mu.Unlock()

	if st.closed {
		return nil
	}

	// Collect and sort remaining rows
	rows := make([]int, 0, len(st.rowBuf))
	for r := range st.rowBuf {
		rows = append(rows, r)
	}
	sort.Ints(rows)

	for _, r := range rows {
		if err := st.flushRowLocked(r); err != nil {
			return err
		}
	}

	st.closed = true
	return st.sw.Flush()
}

// --- Transformer interface implementation ---

func (st *StreamingTransformer) GetCellData(ref CellRef) *CellData {
	return st.reader.GetCellData(ref)
}
func (st *StreamingTransformer) GetCommentedCells() []*CellData {
	return st.reader.GetCommentedCells()
}
func (st *StreamingTransformer) GetFormulaCells() []*CellData {
	return st.reader.GetFormulaCells()
}
func (st *StreamingTransformer) GetSheetNames() []string {
	return st.reader.GetSheetNames()
}
func (st *StreamingTransformer) GetColumnWidth(sheet string, col int) float64 {
	return st.reader.GetColumnWidth(sheet, col)
}
func (st *StreamingTransformer) GetRowHeight(sheet string, row int) float64 {
	return st.reader.GetRowHeight(sheet, row)
}
func (st *StreamingTransformer) GetTargetCellRef(src CellRef) []CellRef {
	return st.reader.GetTargetCellRef(src)
}
func (st *StreamingTransformer) ResetTargetCellRefs() {
	st.reader.ResetTargetCellRefs()
}

// Transform reads from the template and buffers the result for streaming output.
func (st *StreamingTransformer) Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error {
	srcData := st.reader.GetCellData(src)
	if srcData == nil {
		return nil
	}

	// Resolve style
	styleID := 0
	if sid, ok := st.reader.styleCache[src.String()]; ok {
		styleID = sid
	}

	// Handle formula cells
	if srcData.IsFormulaCell() {
		formula := srcData.Formula
		if strings.Contains(formula, ctx.notationBegin) {
			resolved, _, err := ctx.EvaluateCellValue(formula)
			if err == nil && resolved != nil {
				formula = fmt.Sprintf("%v", resolved)
			}
		}
		if err := st.bufferCell(target, nil, styleID, formula); err != nil {
			return err
		}
		srcData.AddTargetPos(target)
		st.reader.addTargetRef(src, target)
		return nil
	}

	// Handle expression cells
	strVal, isStr := srcData.Value.(string)
	if isStr && strings.Contains(strVal, ctx.notationBegin) {
		val, _, err := ctx.EvaluateCellValue(strVal)
		if err != nil {
			return fmt.Errorf("transform cell %s: %w", src, err)
		}

		if hv, ok := val.(HyperlinkValue); ok {
			// Hyperlinks not supported in streaming — write display text only
			if err := st.bufferCell(target, hv.String(), styleID, ""); err != nil {
				return err
			}
		} else {
			if err := st.bufferCell(target, val, styleID, ""); err != nil {
				return err
			}
		}
	} else {
		if err := st.bufferCell(target, srcData.Value, styleID, ""); err != nil {
			return err
		}
	}

	srcData.AddTargetPos(target)
	st.reader.addTargetRef(src, target)
	return nil
}

func (st *StreamingTransformer) ClearCell(ref CellRef) error {
	return st.bufferCell(ref, "", st.lookupStyle(ref), "")
}

func (st *StreamingTransformer) SetFormula(ref CellRef, formula string) error {
	return st.bufferCell(ref, nil, st.lookupStyle(ref), formula)
}

func (st *StreamingTransformer) SetCellValue(ref CellRef, value any) error {
	return st.bufferCell(ref, value, st.lookupStyle(ref), "")
}

// lookupStyle returns the cached style ID for a cell, or 0 if unknown.
func (st *StreamingTransformer) lookupStyle(ref CellRef) int {
	if sid, ok := st.reader.styleCache[ref.String()]; ok {
		return sid
	}
	return 0
}

func (st *StreamingTransformer) SetRowHeight(sheet string, row int, height float64) error {
	// StreamWriter doesn't support per-row height after creation; silently ignore
	return nil
}

func (st *StreamingTransformer) DeleteSheet(name string) error {
	return st.reader.DeleteSheet(name)
}

func (st *StreamingTransformer) SetHidden(name string, hidden bool) error {
	return st.reader.SetHidden(name, hidden)
}

func (st *StreamingTransformer) CopySheet(src, dst string) error {
	return st.reader.CopySheet(src, dst)
}

func (st *StreamingTransformer) AddImage(sheet, cell string, imgBytes []byte, imgType string, scaleX, scaleY float64) error {
	return fmt.Errorf("images are not supported in streaming mode")
}

func (st *StreamingTransformer) MergeCells(sheet, topLeft, bottomRight string) error {
	return st.sw.MergeCell(topLeft, bottomRight)
}

func (st *StreamingTransformer) SetCellHyperLink(ref CellRef, url, display string) error {
	// Hyperlinks not supported in streaming mode — silently ignored
	return nil
}

func (st *StreamingTransformer) SetRecalculateOnOpen(recalc bool) error {
	return st.reader.SetRecalculateOnOpen(recalc)
}

func (st *StreamingTransformer) Write(w io.Writer) error {
	if err := st.Flush(); err != nil {
		return fmt.Errorf("flush streaming rows: %w", err)
	}
	return st.reader.Write(w)
}

func (st *StreamingTransformer) Close() error {
	return st.reader.Close()
}
