package goxls

import (
	"fmt"
	"io"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

// ExcelizeTransformer implements Transformer using excelize.
type ExcelizeTransformer struct {
	file       *excelize.File
	sheets     map[string]*SheetData // in-memory sheet data read from template
	styleCache map[string]int        // "Sheet!A1" → styleID for preservation
	targetRefs map[string][]CellRef  // "Sheet!A1" → list of target positions

	mu sync.Mutex // protects concurrent access
}

// NewExcelizeTransformer creates a Transformer from an excelize file.
func NewExcelizeTransformer(f *excelize.File) (*ExcelizeTransformer, error) {
	tx := &ExcelizeTransformer{
		file:       f,
		sheets:     make(map[string]*SheetData),
		styleCache: make(map[string]int),
		targetRefs: make(map[string][]CellRef),
	}
	if err := tx.readAllCellData(); err != nil {
		return nil, fmt.Errorf("read template data: %w", err)
	}
	return tx, nil
}

// OpenTemplate opens an xlsx file and creates a Transformer.
func OpenTemplate(path string) (*ExcelizeTransformer, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open template %q: %w", path, err)
	}
	return NewExcelizeTransformer(f)
}

// readAllCellData reads all cell data from the template into memory.
func (tx *ExcelizeTransformer) readAllCellData() error {
	for _, sheet := range tx.file.GetSheetList() {
		sd := &SheetData{
			Name:         sheet,
			ColumnWidths: make(map[int]float64),
			Rows:         make(map[int]*RowData),
		}

		// Read column widths
		cols, err := tx.file.GetCols(sheet)
		if err == nil {
			for i := range cols {
				w, err := tx.file.GetColWidth(sheet, ColToName(i))
				if err == nil {
					sd.ColumnWidths[i] = w
				}
			}
		}

		// Read all rows
		rows, err := tx.file.GetRows(sheet)
		if err != nil {
			return fmt.Errorf("read rows from sheet %q: %w", sheet, err)
		}

		for rowIdx, row := range rows {
			rd := &RowData{
				Cells: make(map[int]*CellData),
			}
			h, err := tx.file.GetRowHeight(sheet, rowIdx+1)
			if err == nil {
				rd.Height = h
			}

			for colIdx, cellVal := range row {
				cellName := ColToName(colIdx) + fmt.Sprintf("%d", rowIdx+1)
				ref := NewCellRef(sheet, rowIdx, colIdx)

				cd := &CellData{
					Ref:   ref,
					Value: cellVal,
					Type:  CellString,
				}

				// Detect formula
				formula, err := tx.file.GetCellFormula(sheet, cellName)
				if err == nil && formula != "" {
					cd.Formula = formula
					cd.Type = CellFormula
				}

				// Cache style
				styleID, err := tx.file.GetCellStyle(sheet, cellName)
				if err == nil {
					cd.StyleID = styleID
					tx.styleCache[ref.String()] = styleID
				}

				// Detect cell type from value if not formula
				if cd.Type != CellFormula {
					cd.Type = detectCellType(cellVal)
				}

				rd.Cells[colIdx] = cd
			}

			sd.Rows[rowIdx] = rd
		}

		// Read comments
		comments, err := tx.file.GetComments(sheet)
		if err == nil {
			for _, c := range comments {
				ref, err := ParseCellRef(sheet + "!" + c.Cell)
				if err != nil {
					continue
				}
				// Find or create cell data
				rd, ok := sd.Rows[ref.Row]
				if !ok {
					rd = &RowData{Cells: make(map[int]*CellData)}
					sd.Rows[ref.Row] = rd
				}
				cd, ok := rd.Cells[ref.Col]
				if !ok {
					cd = &CellData{Ref: ref, Type: CellBlank}
					rd.Cells[ref.Col] = cd
				}
				cd.Comment = c.Text
			}
		}

		tx.sheets[sheet] = sd
	}
	return nil
}

func detectCellType(val string) CellType {
	if val == "" {
		return CellBlank
	}
	return CellString
}

// GetCellData returns the cached cell data for the given reference.
func (tx *ExcelizeTransformer) GetCellData(ref CellRef) *CellData {
	sd, ok := tx.sheets[ref.Sheet]
	if !ok {
		return nil
	}
	rd, ok := sd.Rows[ref.Row]
	if !ok {
		return nil
	}
	return rd.Cells[ref.Col]
}

// GetCommentedCells returns all cells that have comments (for template parsing).
func (tx *ExcelizeTransformer) GetCommentedCells() []*CellData {
	var result []*CellData
	for _, sd := range tx.sheets {
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				if cd.Comment != "" {
					result = append(result, cd)
				}
			}
		}
	}
	return result
}

// GetFormulaCells returns all cells that contain formulas.
func (tx *ExcelizeTransformer) GetFormulaCells() []*CellData {
	var result []*CellData
	for _, sd := range tx.sheets {
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				if cd.IsFormulaCell() {
					result = append(result, cd)
				}
			}
		}
	}
	return result
}

// Transform copies a cell from source to target position, evaluating expressions.
func (tx *ExcelizeTransformer) Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error {
	tx.mu.Lock()
	defer tx.mu.Unlock()

	srcData := tx.GetCellData(src)
	if srcData == nil {
		return nil // nothing to transform
	}

	targetSheet := target.Sheet
	if targetSheet == "" {
		targetSheet = src.Sheet
	}
	targetCell := target.CellName()

	// Copy style from source
	if styleID, ok := tx.styleCache[src.String()]; ok {
		tx.file.SetCellStyle(targetSheet, targetCell, targetCell, styleID)
	}

	// Copy column width if source has one
	sd, ok := tx.sheets[src.Sheet]
	if ok {
		if w, ok := sd.ColumnWidths[src.Col]; ok {
			tx.file.SetColWidth(targetSheet, ColToName(target.Col), ColToName(target.Col), w)
		}
	}

	// Copy row height
	if updateRowHeight && ok {
		if rd, ok := sd.Rows[src.Row]; ok && rd.Height > 0 {
			tx.file.SetRowHeight(targetSheet, target.Row+1, rd.Height)
		}
	}

	// Handle formula cells
	if srcData.IsFormulaCell() {
		tx.file.SetCellFormula(targetSheet, targetCell, srcData.Formula)
		srcData.AddTargetPos(target)
		tx.addTargetRef(src, target)
		return nil
	}

	// Handle expression cells
	strVal, isStr := srcData.Value.(string)
	if isStr && strings.Contains(strVal, ctx.notationBegin) {
		val, cellType, err := ctx.EvaluateCellValue(strVal)
		if err != nil {
			return fmt.Errorf("transform cell %s: %w", src, err)
		}
		srcData.EvalResult = val
		srcData.TargetCellType = cellType
		if err := tx.writeTypedValue(targetSheet, targetCell, val, cellType); err != nil {
			return err
		}
	} else {
		// Copy value as-is
		tx.file.SetCellValue(targetSheet, targetCell, srcData.Value)
	}

	srcData.AddTargetPos(target)
	tx.addTargetRef(src, target)
	return nil
}

// writeTypedValue writes a value to a cell with the correct type.
func (tx *ExcelizeTransformer) writeTypedValue(sheet, cell string, value any, cellType CellType) error {
	if value == nil {
		return nil // leave cell blank
	}
	switch cellType {
	case CellFormula:
		return tx.file.SetCellFormula(sheet, cell, fmt.Sprintf("%v", value))
	default:
		return tx.file.SetCellValue(sheet, cell, value)
	}
}

// ClearCell clears a cell's content while preserving style.
func (tx *ExcelizeTransformer) ClearCell(ref CellRef) error {
	tx.mu.Lock()
	defer tx.mu.Unlock()

	sheet := ref.Sheet
	cell := ref.CellName()

	// Preserve style
	styleID, _ := tx.file.GetCellStyle(sheet, cell)
	tx.file.SetCellValue(sheet, cell, "")
	if styleID > 0 {
		tx.file.SetCellStyle(sheet, cell, cell, styleID)
	}
	return nil
}

// SetFormula sets a formula on a cell.
func (tx *ExcelizeTransformer) SetFormula(ref CellRef, formula string) error {
	tx.mu.Lock()
	defer tx.mu.Unlock()
	return tx.file.SetCellFormula(ref.Sheet, ref.CellName(), formula)
}

// SetCellValue sets a value on a cell, preserving style.
func (tx *ExcelizeTransformer) SetCellValue(ref CellRef, value any) error {
	tx.mu.Lock()
	defer tx.mu.Unlock()

	sheet := ref.Sheet
	cell := ref.CellName()
	styleID, _ := tx.file.GetCellStyle(sheet, cell)
	tx.file.SetCellValue(sheet, cell, value)
	if styleID > 0 {
		tx.file.SetCellStyle(sheet, cell, cell, styleID)
	}
	return nil
}

// GetTargetCellRef returns where a source cell was mapped to during transformation.
func (tx *ExcelizeTransformer) GetTargetCellRef(src CellRef) []CellRef {
	return tx.targetRefs[src.String()]
}

// ResetTargetCellRefs clears all source→target mappings.
func (tx *ExcelizeTransformer) ResetTargetCellRefs() {
	tx.targetRefs = make(map[string][]CellRef)
}

func (tx *ExcelizeTransformer) addTargetRef(src, target CellRef) {
	key := src.String()
	tx.targetRefs[key] = append(tx.targetRefs[key], target)
}

// GetSheetNames returns all sheet names.
func (tx *ExcelizeTransformer) GetSheetNames() []string {
	return tx.file.GetSheetList()
}

// GetColumnWidth returns the column width for a sheet/column.
func (tx *ExcelizeTransformer) GetColumnWidth(sheet string, col int) float64 {
	w, err := tx.file.GetColWidth(sheet, ColToName(col))
	if err != nil {
		return 0
	}
	return w
}

// GetRowHeight returns the row height for a sheet/row (0-based row index).
func (tx *ExcelizeTransformer) GetRowHeight(sheet string, row int) float64 {
	h, err := tx.file.GetRowHeight(sheet, row+1)
	if err != nil {
		return 0
	}
	return h
}

// DeleteSheet removes a sheet from the workbook.
func (tx *ExcelizeTransformer) DeleteSheet(name string) error {
	return tx.file.DeleteSheet(name)
}

// SetHidden hides or unhides a sheet.
func (tx *ExcelizeTransformer) SetHidden(name string, hidden bool) error {
	if hidden {
		return tx.file.SetSheetVisible(name, false)
	}
	return tx.file.SetSheetVisible(name, true)
}

// CopySheet copies a sheet to a new name.
func (tx *ExcelizeTransformer) CopySheet(src, dst string) error {
	srcIdx, err := tx.file.GetSheetIndex(src)
	if err != nil {
		return fmt.Errorf("sheet %q not found: %w", src, err)
	}
	newIdx, err := tx.file.NewSheet(dst)
	if err != nil {
		return fmt.Errorf("create sheet %q: %w", dst, err)
	}
	_ = srcIdx
	_ = newIdx
	return tx.file.CopySheet(srcIdx, newIdx)
}

// AddImage inserts an image into a sheet.
func (tx *ExcelizeTransformer) AddImage(sheet string, cell string, imgBytes []byte, imgType string, scaleX, scaleY float64) error {
	tx.mu.Lock()
	defer tx.mu.Unlock()

	ext := ".png"
	switch strings.ToUpper(imgType) {
	case "JPEG", "JPG":
		ext = ".jpg"
	case "GIF":
		ext = ".gif"
	case "BMP":
		ext = ".bmp"
	}

	return tx.file.AddPictureFromBytes(sheet, cell, &excelize.Picture{
		Extension: ext,
		File:      imgBytes,
		Format:    &excelize.GraphicOptions{ScaleX: scaleX, ScaleY: scaleY},
	})
}

// MergeCells merges a cell range.
func (tx *ExcelizeTransformer) MergeCells(sheet, topLeft, bottomRight string) error {
	tx.mu.Lock()
	defer tx.mu.Unlock()
	return tx.file.MergeCell(sheet, topLeft, bottomRight)
}

// Write writes the workbook to the given writer.
func (tx *ExcelizeTransformer) Write(w io.Writer) error {
	return tx.file.Write(w)
}

// Close closes the underlying excelize file.
func (tx *ExcelizeTransformer) Close() error {
	return tx.file.Close()
}

// File returns the underlying excelize file for advanced operations.
func (tx *ExcelizeTransformer) File() *excelize.File {
	return tx.file
}
