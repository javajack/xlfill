package xlfill

// AreaListener is notified before and after each cell transformation.
// Implement this interface to apply conditional styling, logging, or other
// per-cell processing during template expansion.
type AreaListener interface {
	// BeforeTransformCell is called before a cell is transformed from source to target.
	// Return false to skip the default transformation for this cell.
	BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool

	// AfterTransformCell is called after a cell has been transformed.
	AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer)
}

// StyleOverride specifies style modifications to apply to a cell after transformation.
// Nil fields are ignored (no change). Non-nil fields override the cell's current style.
type StyleOverride struct {
	Bold      *bool
	Italic    *bool
	FontColor *string // hex color e.g. "#FF0000"
	FillColor *string // hex background color
	FontSize  *float64
}

// StyleListener is called after cell transformation to optionally modify styling.
// It extends the AreaListener pattern with richer style control.
type StyleListener interface {
	// StyleCell is called after a cell is transformed. Return nil to leave styling unchanged.
	// The value parameter contains the evaluated cell value.
	StyleCell(target CellRef, value any, ctx *Context) *StyleOverride
}
