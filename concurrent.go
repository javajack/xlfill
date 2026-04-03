package xlfill

import (
	"io"
	"sync"
)

// ConcurrentTransformer wraps a Transformer with mutex protection for safe
// concurrent write access. Read operations on immutable template data are
// delegated without synchronization (template data is read-only after init).
type ConcurrentTransformer struct {
	inner Transformer
	mu    sync.Mutex
}

// NewConcurrentTransformer wraps a transformer for concurrent use.
func NewConcurrentTransformer(inner Transformer) *ConcurrentTransformer {
	return &ConcurrentTransformer{inner: inner}
}

// Inner returns the underlying transformer.
func (ct *ConcurrentTransformer) Inner() Transformer { return ct.inner }

// --- Read operations (no lock needed — template data is immutable) ---

func (ct *ConcurrentTransformer) GetCellData(ref CellRef) *CellData {
	return ct.inner.GetCellData(ref)
}
func (ct *ConcurrentTransformer) GetCommentedCells() []*CellData {
	return ct.inner.GetCommentedCells()
}
func (ct *ConcurrentTransformer) GetFormulaCells() []*CellData {
	return ct.inner.GetFormulaCells()
}
func (ct *ConcurrentTransformer) GetSheetNames() []string {
	return ct.inner.GetSheetNames()
}
func (ct *ConcurrentTransformer) GetColumnWidth(sheet string, col int) float64 {
	return ct.inner.GetColumnWidth(sheet, col)
}
func (ct *ConcurrentTransformer) GetRowHeight(sheet string, row int) float64 {
	return ct.inner.GetRowHeight(sheet, row)
}

// --- Write operations (mutex protected) ---

func (ct *ConcurrentTransformer) Transform(src, target CellRef, ctx *Context, updateRowHeight bool) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.Transform(src, target, ctx, updateRowHeight)
}

func (ct *ConcurrentTransformer) ClearCell(ref CellRef) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.ClearCell(ref)
}

func (ct *ConcurrentTransformer) SetFormula(ref CellRef, formula string) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.SetFormula(ref, formula)
}

func (ct *ConcurrentTransformer) SetCellValue(ref CellRef, value any) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.SetCellValue(ref, value)
}

func (ct *ConcurrentTransformer) SetRowHeight(sheet string, row int, height float64) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.SetRowHeight(sheet, row, height)
}

func (ct *ConcurrentTransformer) DeleteSheet(name string) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.DeleteSheet(name)
}

func (ct *ConcurrentTransformer) SetHidden(name string, hidden bool) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.SetHidden(name, hidden)
}

func (ct *ConcurrentTransformer) CopySheet(src, dst string) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.CopySheet(src, dst)
}

func (ct *ConcurrentTransformer) AddImage(sheet, cell string, imgBytes []byte, imgType string, scaleX, scaleY float64) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.AddImage(sheet, cell, imgBytes, imgType, scaleX, scaleY)
}

func (ct *ConcurrentTransformer) MergeCells(sheet, topLeft, bottomRight string) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.MergeCells(sheet, topLeft, bottomRight)
}

func (ct *ConcurrentTransformer) SetCellHyperLink(ref CellRef, url, display string) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.SetCellHyperLink(ref, url, display)
}

func (ct *ConcurrentTransformer) SetRecalculateOnOpen(recalc bool) error {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.SetRecalculateOnOpen(recalc)
}

// --- Target tracking (synchronized because writes happen concurrently) ---

func (ct *ConcurrentTransformer) GetTargetCellRef(src CellRef) []CellRef {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	return ct.inner.GetTargetCellRef(src)
}

func (ct *ConcurrentTransformer) ResetTargetCellRefs() {
	ct.mu.Lock()
	defer ct.mu.Unlock()
	ct.inner.ResetTargetCellRefs()
}

// --- I/O (not concurrent — called after processing) ---

func (ct *ConcurrentTransformer) Write(w io.Writer) error {
	return ct.inner.Write(w)
}

func (ct *ConcurrentTransformer) Close() error {
	return ct.inner.Close()
}
