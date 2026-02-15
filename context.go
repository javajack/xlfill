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

	// Cached merged map for expression evaluation.
	// Invalidated (set to nil) whenever runVars change.
	cachedMap map[string]any
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
	}
	for _, opt := range opts {
		opt(c)
	}
	return c
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
	c.invalidateCache()
}

// RemoveVar removes a variable from the data map.
func (c *Context) RemoveVar(name string) {
	delete(c.data, name)
	c.invalidateCache()
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
// The result is cached and reused until runVars are modified.
func (c *Context) ToMap() map[string]any {
	if c.cachedMap != nil {
		return c.cachedMap
	}
	m := make(map[string]any, len(c.data)+len(c.runVars)+2)
	for k, v := range c.data {
		m[k] = v
	}
	for k, v := range c.runVars {
		m[k] = v
	}
	// Built-in functions
	if _, ok := m["hyperlink"]; !ok {
		m["hyperlink"] = Hyperlink
	}
	c.cachedMap = m
	return m
}

// invalidateCache clears the cached merged map.
func (c *Context) invalidateCache() {
	c.cachedMap = nil
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
func (c *Context) setRunVar(name string, value any) {
	c.runVars[name] = value
	c.invalidateCache()
}

// removeRunVar removes a run variable.
func (c *Context) removeRunVar(name string) {
	delete(c.runVars, name)
	c.invalidateCache()
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
