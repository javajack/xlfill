package xlfill

import (
	"fmt"
	"strings"
	"sync"

	"github.com/expr-lang/expr"
	"github.com/expr-lang/expr/vm"
)

// ExpressionEvaluator evaluates template expressions.
type ExpressionEvaluator interface {
	Evaluate(expression string, data map[string]any) (any, error)
	IsConditionTrue(condition string, data map[string]any) (bool, error)
}

// exprEvaluator implements ExpressionEvaluator using expr-lang/expr.
type exprEvaluator struct {
	cache sync.Map // expression string → compiled *vm.Program
}

// NewExpressionEvaluator creates a new expression evaluator backed by expr-lang/expr.
func NewExpressionEvaluator() ExpressionEvaluator {
	return &exprEvaluator{}
}

func (e *exprEvaluator) Evaluate(expression string, data map[string]any) (any, error) {
	if expression == "" {
		return nil, nil
	}
	program, err := e.compile(expression, data)
	if err != nil {
		return nil, fmt.Errorf("compile expression %q: %w", expression, err)
	}
	result, err := expr.Run(program, data)
	if err != nil {
		return nil, fmt.Errorf("evaluate expression %q: %w", expression, err)
	}
	return result, nil
}

func (e *exprEvaluator) IsConditionTrue(condition string, data map[string]any) (bool, error) {
	result, err := e.Evaluate(condition, data)
	if err != nil {
		return false, err
	}
	if result == nil {
		return false, nil // nil treated as false (JXLS v3 behavior)
	}
	b, ok := result.(bool)
	if !ok {
		return false, fmt.Errorf("condition %q evaluated to %T, expected bool", condition, result)
	}
	return b, nil
}

func (e *exprEvaluator) compile(expression string, env map[string]any) (*vm.Program, error) {
	if cached, ok := e.cache.Load(expression); ok {
		return cached.(*vm.Program), nil
	}
	program, err := expr.Compile(expression, expr.Env(env), expr.AllowUndefinedVariables())
	if err != nil {
		return nil, err
	}
	e.cache.Store(expression, program)
	return program, nil
}

// ExpressionSegment represents a part of a cell value: either literal text or an expression.
type ExpressionSegment struct {
	IsExpression bool
	Text         string // literal text or expression content (without delimiters)
}

// ParseExpressions splits a cell value into segments of literal text and expressions.
// For example, "Name: ${e.Name}" → [{false, "Name: "}, {true, "e.Name"}]
func ParseExpressions(value string, begin, end string) []ExpressionSegment {
	if begin == "" || end == "" {
		begin = "${"
		end = "}"
	}

	var segments []ExpressionSegment
	remaining := value

	for {
		startIdx := strings.Index(remaining, begin)
		if startIdx < 0 {
			break
		}

		// Find matching end delimiter, accounting for nested braces
		searchFrom := startIdx + len(begin)
		endIdx := findMatchingEnd(remaining[searchFrom:], begin, end)
		if endIdx < 0 {
			break
		}
		endIdx += searchFrom

		// Add literal text before expression
		if startIdx > 0 {
			segments = append(segments, ExpressionSegment{
				IsExpression: false,
				Text:         remaining[:startIdx],
			})
		}

		// Add expression
		exprText := remaining[startIdx+len(begin) : endIdx]
		segments = append(segments, ExpressionSegment{
			IsExpression: true,
			Text:         exprText,
		})

		remaining = remaining[endIdx+len(end):]
	}

	// Add remaining literal text
	if remaining != "" {
		segments = append(segments, ExpressionSegment{
			IsExpression: false,
			Text:         remaining,
		})
	}

	return segments
}

// findMatchingEnd finds the position of the matching end delimiter,
// handling nested begin/end pairs.
func findMatchingEnd(s string, begin, end string) int {
	depth := 0
	for i := 0; i <= len(s)-len(end); i++ {
		if strings.HasPrefix(s[i:], begin) {
			depth++
		} else if strings.HasPrefix(s[i:], end) {
			if depth == 0 {
				return i
			}
			depth--
		}
	}
	return -1
}

// IsExpressionOnly returns true if the value is a single expression with no surrounding text.
// e.g., "${e.Name}" is expression-only, but "Name: ${e.Name}" is not.
func IsExpressionOnly(value string, begin, end string) bool {
	if begin == "" || end == "" {
		begin = "${"
		end = "}"
	}
	trimmed := strings.TrimSpace(value)
	if !strings.HasPrefix(trimmed, begin) || !strings.HasSuffix(trimmed, end) {
		return false
	}
	// Check there's exactly one expression and no other content
	inner := trimmed[len(begin) : len(trimmed)-len(end)]
	return !strings.Contains(inner, begin)
}

// ExtractSingleExpression extracts the expression from a value like "${e.Name}".
// Returns the expression string and true if it's a single expression, or ("", false) otherwise.
func ExtractSingleExpression(value string, begin, end string) (string, bool) {
	if begin == "" || end == "" {
		begin = "${"
		end = "}"
	}
	trimmed := strings.TrimSpace(value)
	if !strings.HasPrefix(trimmed, begin) || !strings.HasSuffix(trimmed, end) {
		return "", false
	}
	inner := trimmed[len(begin) : len(trimmed)-len(end)]
	if strings.Contains(inner, begin) {
		return "", false
	}
	return inner, true
}
