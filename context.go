package xlfill

import (
	"fmt"
	"strings"
)

// Context holds template data and provides expression evaluation.
// It manages both user-provided data and loop iteration variables (runVars).
type Context struct {
	data           map[string]any
	runVars        map[string]any
	evaluator      ExpressionEvaluator
	notationBegin  string
	notationEnd    string
	updateCellData bool
	clearCells     bool

	// Differential map caching: instead of rebuilding the merged map on every
	// runVar change, we keep the map alive and apply in-place updates.
	cachedMap    map[string]any
	mapDirtyKeys map[string]struct{} // runVar keys that changed since last ToMap()
	mapNeedsFull bool                // true when data changed (requires full rebuild)

	// Deferred actions registered during command processing.
	deferred *DeferredRegistry

	// Custom template functions (from WithFunction API).
	customFunctions map[string]any

	// i18n resource bundle for translations.
	i18nBundle map[string]string
}

// ContextOption configures a Context.
type ContextOption func(*Context)

// WithNotation sets custom expression notation delimiters.
func WithNotation(begin, end string) ContextOption {
	return func(c *Context) {
		c.notationBegin = begin
		c.notationEnd = end
	}
}

// WithEvaluator sets a custom expression evaluator.
func WithEvaluator(ev ExpressionEvaluator) ContextOption {
	return func(c *Context) {
		c.evaluator = ev
	}
}

// WithUpdateCellData enables/disables cell data tracking for formulas.
func WithUpdateCellData(enabled bool) ContextOption {
	return func(c *Context) {
		c.updateCellData = enabled
	}
}

// WithClearCells enables/disables clearing of template cells after processing.
func WithClearCells(enabled bool) ContextOption {
	return func(c *Context) {
		c.clearCells = enabled
	}
}

// WithCustomFunctions sets custom template functions on the context.
func WithCustomFunctions(fns map[string]any) ContextOption {
	return func(c *Context) {
		c.customFunctions = fns
	}
}

// WithI18nBundle sets the i18n resource bundle for the t() template function.
func WithI18nBundle(bundle map[string]string) ContextOption {
	return func(c *Context) {
		c.i18nBundle = bundle
	}
}

// NewContext creates a new Context with the given data and options.
func NewContext(data map[string]any, opts ...ContextOption) *Context {
	if data == nil {
		data = make(map[string]any)
	}
	c := &Context{
		data:           data,
		runVars:        make(map[string]any),
		evaluator:      NewExpressionEvaluator(),
		notationBegin:  "${",
		notationEnd:    "}",
		updateCellData: true,
		clearCells:     true,
		mapDirtyKeys:   make(map[string]struct{}, 4),
		mapNeedsFull:   true, // first ToMap() must do full build
		deferred:       NewDeferredRegistry(),
	}
	for _, opt := range opts {
		opt(c)
	}
	return c
}

// RegisterDeferred registers a deferred action to be executed after all areas are processed.
func (c *Context) RegisterDeferred(action DeferredAction) {
	c.deferred.Add(action)
}

// Deferred returns the deferred registry for this context.
func (c *Context) Deferred() *DeferredRegistry {
	return c.deferred
}

// Clone creates an independent copy of the Context suitable for parallel processing.
// The data map is shared (read-only), but runVars and the cached map are independent.
// The evaluator is shared (its sync.Map cache is already thread-safe).
// The deferred registry is shared (it is thread-safe via mutex).
func (c *Context) Clone() *Context {
	return &Context{
		data:            c.data, // shared read-only
		runVars:         make(map[string]any),
		evaluator:       c.evaluator, // thread-safe
		notationBegin:   c.notationBegin,
		notationEnd:     c.notationEnd,
		updateCellData:  c.updateCellData,
		clearCells:      c.clearCells,
		mapDirtyKeys:    make(map[string]struct{}, 4),
		mapNeedsFull:    true,
		deferred:        c.deferred,        // shared, thread-safe
		customFunctions: c.customFunctions,  // shared read-only
		i18nBundle:      c.i18nBundle,       // shared read-only
	}
}

// GetVar returns a variable value. Checks runVars first, then data.
func (c *Context) GetVar(name string) any {
	if v, ok := c.runVars[name]; ok {
		return v
	}
	return c.data[name]
}

// PutVar sets a variable in the data map.
func (c *Context) PutVar(name string, value any) {
	c.data[name] = value
	c.mapNeedsFull = true // data changed, need full rebuild
}

// RemoveVar removes a variable from the data map.
func (c *Context) RemoveVar(name string) {
	delete(c.data, name)
	c.mapNeedsFull = true
}

// ContainsVar returns true if the variable exists in either runVars or data.
func (c *Context) ContainsVar(name string) bool {
	if _, ok := c.runVars[name]; ok {
		return true
	}
	_, ok := c.data[name]
	return ok
}

// ToMap returns a merged map of data and runVars. RunVars override data.
// Built-in functions are always available.
// Uses differential updates: when only runVars change, the existing map is
// updated in-place rather than rebuilt from scratch.
func (c *Context) ToMap() map[string]any {
	if c.cachedMap != nil && !c.mapNeedsFull && len(c.mapDirtyKeys) == 0 {
		return c.cachedMap
	}

	if c.cachedMap == nil || c.mapNeedsFull {
		// Full rebuild
		m := make(map[string]any, len(c.data)+len(c.runVars)+16)
		for k, v := range c.data {
			m[k] = v
		}
		for k, v := range c.runVars {
			m[k] = v
		}
		// Register built-in functions (user data takes precedence)
		if _, ok := m["hyperlink"]; !ok {
			m["hyperlink"] = Hyperlink
		}
		if _, ok := m["comment"]; !ok {
			m["comment"] = Comment
		}
		registerBuiltins(m, c.i18nBundle)
		// Merge custom functions (user-provided via WithFunction)
		for k, v := range c.customFunctions {
			if _, ok := m[k]; !ok {
				m[k] = v
			}
		}
		c.cachedMap = m
		c.mapNeedsFull = false
		clear(c.mapDirtyKeys)
		return m
	}

	// Differential update: only apply changed runVars (deduplicated via map keys)
	for key := range c.mapDirtyKeys {
		if val, ok := c.runVars[key]; ok {
			c.cachedMap[key] = val
		} else {
			// Key was removed from runVars — restore from data, builtin, custom, or delete
			if val, ok := c.data[key]; ok {
				c.cachedMap[key] = val
			} else if key == "hyperlink" {
				c.cachedMap[key] = Hyperlink
			} else if key == "comment" {
				c.cachedMap[key] = Comment
			} else if fn, ok := c.customFunctions[key]; ok {
				c.cachedMap[key] = fn
			} else if restoreBuiltin(c.cachedMap, key, c.i18nBundle) {
				// restored by restoreBuiltin
			} else {
				delete(c.cachedMap, key)
			}
		}
	}
	clear(c.mapDirtyKeys)
	return c.cachedMap
}

// Evaluate evaluates an expression string using the merged data.
func (c *Context) Evaluate(expression string) (any, error) {
	return c.evaluator.Evaluate(expression, c.ToMap())
}

// IsConditionTrue evaluates a boolean condition.
func (c *Context) IsConditionTrue(condition string) (bool, error) {
	return c.evaluator.IsConditionTrue(condition, c.ToMap())
}

// EvaluateCellValue evaluates a cell value string, processing embedded expressions.
// If the value is a single expression like "${e.Name}", the result is typed (number, bool, etc.).
// If mixed content like "Name: ${e.Name}", the result is always a string.
func (c *Context) EvaluateCellValue(value string) (any, CellType, error) {
	// Check if it's a single expression
	exprStr, isSingle := ExtractSingleExpression(value, c.notationBegin, c.notationEnd)
	if isSingle {
		result, err := c.Evaluate(exprStr)
		if err != nil {
			return nil, CellBlank, fmt.Errorf("evaluate %q: %w", value, err)
		}
		return result, inferCellType(result), nil
	}

	// Parse and evaluate all expressions in mixed content
	segments := ParseExpressions(value, c.notationBegin, c.notationEnd)
	if len(segments) == 0 {
		return value, CellString, nil
	}

	// Check if there are any expressions at all
	hasExpr := false
	for _, seg := range segments {
		if seg.IsExpression {
			hasExpr = true
			break
		}
	}
	if !hasExpr {
		return value, CellString, nil
	}

	// Build result string
	var b strings.Builder
	for _, seg := range segments {
		if seg.IsExpression {
			val, err := c.Evaluate(seg.Text)
			if err != nil {
				return nil, CellBlank, fmt.Errorf("evaluate expression %q in %q: %w", seg.Text, value, err)
			}
			if val != nil {
				fmt.Fprintf(&b, "%v", val)
			}
		} else {
			b.WriteString(seg.Text)
		}
	}
	return b.String(), CellString, nil
}

// inferCellType determines the CellType from a Go value.
func inferCellType(v any) CellType {
	if v == nil {
		return CellBlank
	}
	switch v.(type) {
	case bool:
		return CellBoolean
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		return CellNumber
	case string:
		return CellString
	default:
		return CellString
	}
}

// setRunVar sets a run variable (loop iteration variable).
// Uses differential cache update instead of full invalidation.
func (c *Context) setRunVar(name string, value any) {
	c.runVars[name] = value
	c.mapDirtyKeys[name] = struct{}{}
}

// removeRunVar removes a run variable.
func (c *Context) removeRunVar(name string) {
	delete(c.runVars, name)
	c.mapDirtyKeys[name] = struct{}{}
}

// RunVar manages scoped loop variables with automatic save/restore.
// Use with defer: rv := NewRunVar(ctx, "e"); defer rv.Close()
type RunVar struct {
	ctx      *Context
	varName  string
	oldValue any
	hadOld   bool
	idxName  string
	oldIdx   any
	hadIdx   bool
}

// NewRunVar creates a new RunVar for a single loop variable.
func NewRunVar(ctx *Context, varName string) *RunVar {
	rv := &RunVar{
		ctx:     ctx,
		varName: varName,
	}
	if old, ok := ctx.runVars[varName]; ok {
		rv.oldValue = old
		rv.hadOld = true
	}
	return rv
}

// NewRunVarWithIndex creates a RunVar for a loop variable and its index.
func NewRunVarWithIndex(ctx *Context, varName, idxName string) *RunVar {
	rv := NewRunVar(ctx, varName)
	rv.idxName = idxName
	if old, ok := ctx.runVars[idxName]; ok {
		rv.oldIdx = old
		rv.hadIdx = true
	}
	return rv
}

// Set sets the loop variable value.
func (rv *RunVar) Set(value any) {
	rv.ctx.setRunVar(rv.varName, value)
}

// SetWithIndex sets both the loop variable and its index.
func (rv *RunVar) SetWithIndex(value any, index int) {
	rv.ctx.setRunVar(rv.varName, value)
	if rv.idxName != "" {
		rv.ctx.setRunVar(rv.idxName, index)
	}
}

// Close restores the previous variable values. Designed for use with defer.
func (rv *RunVar) Close() {
	if rv.hadOld {
		rv.ctx.setRunVar(rv.varName, rv.oldValue)
	} else {
		rv.ctx.removeRunVar(rv.varName)
	}
	if rv.idxName != "" {
		if rv.hadIdx {
			rv.ctx.setRunVar(rv.idxName, rv.oldIdx)
		} else {
			rv.ctx.removeRunVar(rv.idxName)
		}
	}
}
