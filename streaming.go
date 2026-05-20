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
// Multi-sheet streaming: a single StreamingTransformer can stream multiple
// sheets at once, each backed by its own StreamWriter and row buffer. Writes
// to sheets that are NOT in the streamed set are delegated to the underlying
// ExcelizeTransformer (so hyperlinks, images, and formula remapping still work
// on those sheets).
//
// Limitations on streamed sheets:
//   - Formula post-processing (reference remapping) is not supported
//   - Hyperlinks are silently dropped and surfaced via Filler.Warnings()
//   - Images return an error
//   - Per-row height changes after creation are not supported
//   - Rows are written in ascending order (guaranteed by area processing)
type StreamingTransformer struct {
	reader        *ExcelizeTransformer
	sws           map[string]*excelize.StreamWriter
	rowBufs       map[string]map[int]map[int]*streamCell
	nextFlushRows map[string]int
	mu            sync.Mutex
	closed        bool

	// Warnings collected during streaming (e.g. dropped hyperlinks).
	// Inspect via Warnings() — populated by Transform.
	warnMu   sync.Mutex
	warnings []string
}

// NewStreamingTransformer creates a streaming transformer that streams a single sheet.
// Equivalent to NewStreamingTransformerForSheets(reader, []string{sheet}).
func NewStreamingTransformer(reader *ExcelizeTransformer, sheet string) (*StreamingTransformer, error) {
	return NewStreamingTransformerForSheets(reader, []string{sheet})
}

// NewStreamingTransformerForSheets creates a streaming transformer that streams
// every sheet in the given list. Writes to other sheets are delegated to the
// underlying ExcelizeTransformer.
//
// Pass nil or an empty list to stream every sheet in the workbook.
//
// Sheets named in the list that do not exist in the workbook are skipped with
// a warning (inspect via Warnings()). If no listed sheet exists at all, the
// returned transformer streams nothing and behaves like a passthrough to the
// underlying ExcelizeTransformer.
func NewStreamingTransformerForSheets(reader *ExcelizeTransformer, sheets []string) (*StreamingTransformer, error) {
	if len(sheets) == 0 {
		sheets = reader.GetSheetNames()
	}
	existing := make(map[string]struct{}, len(reader.GetSheetNames()))
	for _, s := range reader.GetSheetNames() {
		existing[s] = struct{}{}
	}
	st := &StreamingTransformer{
		reader:        reader,
		sws:           make(map[string]*excelize.StreamWriter, len(sheets)),
		rowBufs:       make(map[string]map[int]map[int]*streamCell, len(sheets)),
		nextFlushRows: make(map[string]int, len(sheets)),
	}
	for _, sheet := range sheets {
		if _, ok := existing[sheet]; !ok {
			st.addWarning(fmt.Sprintf("streaming sheet %q not found in template — skipped", sheet))
			continue
		}
		sw, err := reader.file.NewStreamWriter(sheet)
		if err != nil {
			return nil, fmt.Errorf("create stream writer for sheet %q: %w", sheet, err)
		}
		st.sws[sheet] = sw
		st.rowBufs[sheet] = make(map[int]map[int]*streamCell)

		if sd, ok := reader.sheets[sheet]; ok {
			for col, w := range sd.ColumnWidths {
				colNum := col + 1
				sw.SetColWidth(colNum, colNum, w)
			}
		}
	}
	return st, nil
}

// isStreamedSheet reports whether writes to the named sheet should be streamed.
func (st *StreamingTransformer) isStreamedSheet(sheet string) bool {
	_, ok := st.sws[sheet]
	return ok
}

// StreamedSheets returns the sorted list of sheet names that are being streamed.
func (st *StreamingTransformer) StreamedSheets() []string {
	names := make([]string, 0, len(st.sws))
	for name := range st.sws {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}

// Warnings returns warnings collected during streaming
// (e.g. hyperlinks dropped because they can't be expressed via StreamWriter).
// Safe to call concurrently.
func (st *StreamingTransformer) Warnings() []string {
	st.warnMu.Lock()
	defer st.warnMu.Unlock()
	out := make([]string, len(st.warnings))
	copy(out, st.warnings)
	return out
}

func (st *StreamingTransformer) addWarning(msg string) {
	st.warnMu.Lock()
	st.warnings = append(st.warnings, msg)
	st.warnMu.Unlock()
}

// bufferCell stores a cell value for later flushing. Caller must ensure
// ref.Sheet is in the streamed set (use isStreamedSheet to check).
func (st *StreamingTransformer) bufferCell(ref CellRef, value any, styleID int, formula string) error {
	st.mu.Lock()
	defer st.mu.Unlock()

	row := ref.Row
	sheet := ref.Sheet

	rowBuf := st.rowBufs[sheet]
	nextFlushRow := st.nextFlushRows[sheet]

	// Flush all completed rows below this one
	for r := nextFlushRow; r < row; r++ {
		if err := st.flushRowLocked(sheet, r); err != nil {
			return err
		}
	}

	if rowBuf[row] == nil {
		rowBuf[row] = make(map[int]*streamCell)
	}
	rowBuf[row][ref.Col] = &streamCell{value: value, styleID: styleID, formula: formula}
	return nil
}

// flushRowLocked writes a buffered row to the StreamWriter for the given sheet.
// Caller must hold mu.
func (st *StreamingTransformer) flushRowLocked(sheet string, row int) error {
	rowBuf := st.rowBufs[sheet]
	sw := st.sws[sheet]
	cells := rowBuf[row]
	if len(cells) == 0 {
		st.nextFlushRows[sheet] = row + 1
		return nil
	}

	minCol, maxCol := int(^uint(0)>>1), 0
	for col := range cells {
		if col < minCol {
			minCol = col
		}
		if col > maxCol {
			maxCol = col
		}
	}
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
	if err := sw.SetRow(startCell, values); err != nil {
		return fmt.Errorf("stream write row %d on sheet %q: %w", row+1, sheet, err)
	}

	delete(rowBuf, row)
	st.nextFlushRows[sheet] = row + 1
	return nil
}

// Flush writes all remaining buffered rows on every streamed sheet and
// finalizes each StreamWriter.
func (st *StreamingTransformer) Flush() error {
	st.mu.Lock()
	defer st.mu.Unlock()

	if st.closed {
		return nil
	}

	// Sort sheet names for deterministic flush order.
	sheets := make([]string, 0, len(st.sws))
	for s := range st.sws {
		sheets = append(sheets, s)
	}
	sort.Strings(sheets)

	for _, sheet := range sheets {
		rowBuf := st.rowBufs[sheet]
		rows := make([]int, 0, len(rowBuf))
		for r := range rowBuf {
			rows = append(rows, r)
		}
		sort.Ints(rows)
		for _, r := range rows {
			if err := st.flushRowLocked(sheet, r); err != nil {
				return err
			}
		}
		if err := st.sws[sheet].Flush(); err != nil {
			return fmt.Errorf("flush stream writer for sheet %q: %w", sheet, err)
		}
	}

	st.closed = true
	return nil
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

// Transform routes to the streaming buffer for streamed sheets, or to the
// underlying ExcelizeTransformer for non-streamed sheets (preserving full
// hyperlink and image support there).
func (st *StreamingTransformer) Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error {
	if !st.isStreamedSheet(target.Sheet) {
		return st.reader.Transform(src, target, ctx, updateRowHeight)
	}

	srcData := st.reader.GetCellData(src)
	if srcData == nil {
		return nil
	}

	styleID := 0
	if sid, ok := st.reader.styleCache[src.String()]; ok {
		styleID = sid
	}

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

	strVal, isStr := srcData.Value.(string)
	if isStr && strings.Contains(strVal, ctx.notationBegin) {
		val, _, err := ctx.EvaluateCellValue(strVal)
		if err != nil {
			return fmt.Errorf("transform cell %s: %w", src, err)
		}

		if hv, ok := val.(HyperlinkValue); ok {
			// Streaming can't write hyperlinks. Record a warning and fall
			// back to the display text so users see the link target text.
			st.addWarning(fmt.Sprintf("hyperlink dropped at %s (URL=%q) — streaming mode cannot write hyperlinks", target, hv.URL))
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
	if !st.isStreamedSheet(ref.Sheet) {
		return st.reader.ClearCell(ref)
	}
	return st.bufferCell(ref, "", st.lookupStyle(ref), "")
}

func (st *StreamingTransformer) SetFormula(ref CellRef, formula string) error {
	if !st.isStreamedSheet(ref.Sheet) {
		return st.reader.SetFormula(ref, formula)
	}
	return st.bufferCell(ref, nil, st.lookupStyle(ref), formula)
}

func (st *StreamingTransformer) SetCellValue(ref CellRef, value any) error {
	if !st.isStreamedSheet(ref.Sheet) {
		return st.reader.SetCellValue(ref, value)
	}
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
	if !st.isStreamedSheet(sheet) {
		return st.reader.SetRowHeight(sheet, row, height)
	}
	// StreamWriter doesn't support per-row height after creation; silently ignore.
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
	if !st.isStreamedSheet(sheet) {
		return st.reader.AddImage(sheet, cell, imgBytes, imgType, scaleX, scaleY)
	}
	return fmt.Errorf("images are not supported in streaming mode (sheet %q)", sheet)
}

func (st *StreamingTransformer) MergeCells(sheet, topLeft, bottomRight string) error {
	if !st.isStreamedSheet(sheet) {
		return st.reader.MergeCells(sheet, topLeft, bottomRight)
	}
	return st.sws[sheet].MergeCell(topLeft, bottomRight)
}

func (st *StreamingTransformer) SetCellHyperLink(ref CellRef, url, display string) error {
	if !st.isStreamedSheet(ref.Sheet) {
		return st.reader.SetCellHyperLink(ref, url, display)
	}
	st.addWarning(fmt.Sprintf("hyperlink dropped at %s (URL=%q) — streaming mode cannot write hyperlinks", ref, url))
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
