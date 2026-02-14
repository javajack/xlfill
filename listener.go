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
