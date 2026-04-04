package xlfill

import (
	"fmt"
	"io"
	"strings"
	"sync"
	"sync/atomic"
	"time"
)

// DebugTracer writes structured trace output during template processing.
// It implements AreaListener to hook into the transform pipeline.
// All methods are safe for concurrent use.
type DebugTracer struct {
	mu        sync.Mutex
	w         io.Writer
	indent    int
	startTime time.Time
	cellCount atomic.Int64
}

// NewDebugTracer creates a DebugTracer writing to w.
func NewDebugTracer(w io.Writer) *DebugTracer {
	return &DebugTracer{w: w, startTime: time.Now()}
}

func (d *DebugTracer) prefix() string {
	d.mu.Lock()
	n := d.indent
	d.mu.Unlock()
	return strings.Repeat("  ", n)
}

// TraceArea logs area processing start.
func (d *DebugTracer) TraceArea(area *Area, target CellRef) {
	lastCell := NewCellRef(
		area.StartCell.Sheet,
		area.StartCell.Row+area.AreaSize.Height-1,
		area.StartCell.Col+area.AreaSize.Width-1,
	)
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[area] %s:%s %s -> target %s, %d bindings\n",
		strings.Repeat("  ", d.indent), area.StartCell, lastCell.CellName(), area.AreaSize, target, len(area.Bindings))
	d.mu.Unlock()
}

// TraceCommand logs command execution.
func (d *DebugTracer) TraceCommand(cmd Command, target CellRef, attrs string) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[%s] %s%s\n", strings.Repeat("  ", d.indent), cmd.Name(), target, attrs)
	d.mu.Unlock()
}

// TraceEachStart logs each command iteration start.
func (d *DebugTracer) TraceEachStart(items int, varName, direction string) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[each] %d items, var=%q, direction=%s\n",
		strings.Repeat("  ", d.indent), items, varName, direction)
	d.mu.Unlock()
}

// TraceIteration logs a single loop iteration.
func (d *DebugTracer) TraceIteration(index int, target CellRef) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[iter %d] -> %s\n", strings.Repeat("  ", d.indent), index, target)
	d.mu.Unlock()
}

// TraceExpr logs expression evaluation.
func (d *DebugTracer) TraceExpr(expr string, result any, elapsed time.Duration) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[expr] ${%s} -> %v (%s)\n",
		strings.Repeat("  ", d.indent), expr, result, elapsed.Round(time.Microsecond))
	d.mu.Unlock()
}

// TraceRepeat logs repeat command execution.
func (d *DebugTracer) TraceRepeat(count int, varName string) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[repeat] count=%d, var=%q\n", strings.Repeat("  ", d.indent), count, varName)
	d.mu.Unlock()
}

// TraceIf logs if command evaluation.
func (d *DebugTracer) TraceIf(condition string, result bool) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[if] condition=%q -> %v\n", strings.Repeat("  ", d.indent), condition, result)
	d.mu.Unlock()
}

// TraceDeferredAction logs execution of a deferred action.
func (d *DebugTracer) TraceDeferredAction(name, sheet string, startRow, endRow int) {
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[deferred] %s on %s rows %d-%d\n",
		strings.Repeat("  ", d.indent), name, sheet, startRow+1, endRow+1)
	d.mu.Unlock()
}

// TraceDone logs processing completion.
func (d *DebugTracer) TraceDone() {
	elapsed := time.Since(d.startTime)
	count := d.cellCount.Load()
	d.mu.Lock()
	fmt.Fprintf(d.w, "%s[done] %d cells transformed, %s total\n",
		strings.Repeat("  ", d.indent), count, elapsed.Round(time.Microsecond))
	d.mu.Unlock()
}

// Indent increases trace indentation.
func (d *DebugTracer) Indent() {
	d.mu.Lock()
	d.indent++
	d.mu.Unlock()
}

// Dedent decreases trace indentation.
func (d *DebugTracer) Dedent() {
	d.mu.Lock()
	if d.indent > 0 {
		d.indent--
	}
	d.mu.Unlock()
}

// BeforeTransformCell implements AreaListener.
func (d *DebugTracer) BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool {
	d.cellCount.Add(1)
	return true // proceed with default transform
}

// AfterTransformCell implements AreaListener.
func (d *DebugTracer) AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer) {}
