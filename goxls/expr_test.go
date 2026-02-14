package goxls

import (
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

type testEmployee struct {
	Name    string
	Age     int
	Payment float64
	Active  bool
	Address *testAddress
}

type testAddress struct {
	City    string
	Country string
}

func newTestEvalEnv() map[string]any {
	return map[string]any{
		"e": testEmployee{
			Name:    "Alice",
			Age:     30,
			Payment: 5000.0,
			Active:  true,
			Address: &testAddress{City: "London", Country: "UK"},
		},
		"list":  []int{1, 2, 3},
		"empty": []int{},
	}
}

// --- ExpressionEvaluator Tests ---

func TestExpr_SimpleProperty(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("e.Name", newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, "Alice", result)
}

func TestExpr_NestedProperty(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("e.Address.City", newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, "London", result)
}

func TestExpr_MapAccess(t *testing.T) {
	ev := NewExpressionEvaluator()
	data := map[string]any{
		"data": map[string]any{"key": "value"},
	}
	result, err := ev.Evaluate(`data["key"]`, data)
	require.NoError(t, err)
	assert.Equal(t, "value", result)
}

func TestExpr_Arithmetic(t *testing.T) {
	ev := NewExpressionEvaluator()
	data := map[string]any{"a": 10, "b": 20}
	result, err := ev.Evaluate("a + b", data)
	require.NoError(t, err)
	assert.Equal(t, 30, result)
}

func TestExpr_Comparison(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("e.Payment > 2000", newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, true, result)
}

func TestExpr_LogicalAnd(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("e.Age > 0 && e.Payment < 10000", newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, true, result)
}

func TestExpr_LogicalOr(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("e.Age > 100 || e.Payment > 1000", newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, true, result)
}

func TestExpr_Ternary(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate(`e.Active ? "Yes" : "No"`, newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, "Yes", result)
}

func TestExpr_LenFunction(t *testing.T) {
	ev := NewExpressionEvaluator()
	env := newTestEvalEnv()

	result, err := ev.Evaluate("len(list)", env)
	require.NoError(t, err)
	assert.Equal(t, 3, result)

	result, err = ev.Evaluate("len(empty) == 0", env)
	require.NoError(t, err)
	assert.Equal(t, true, result)
}

func TestExpr_NullVariable(t *testing.T) {
	ev := NewExpressionEvaluator()
	env := map[string]any{"x": nil}
	result, err := ev.Evaluate("x", env)
	require.NoError(t, err)
	assert.Nil(t, result)
}

func TestExpr_StringConcat(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate(`"Hello " + e.Name`, newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, "Hello Alice", result)
}

func TestExpr_ErrorExpression(t *testing.T) {
	ev := NewExpressionEvaluator()
	_, err := ev.Evaluate("!!!invalid!!!", map[string]any{})
	assert.Error(t, err)
}

func TestExpr_IsConditionTrue(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.IsConditionTrue("e.Payment > 2000", newTestEvalEnv())
	require.NoError(t, err)
	assert.True(t, result)
}

func TestExpr_IsConditionFalse(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.IsConditionTrue("e.Payment > 99999", newTestEvalEnv())
	require.NoError(t, err)
	assert.False(t, result)
}

func TestExpr_NilCondition(t *testing.T) {
	ev := NewExpressionEvaluator()
	// Evaluating a nil variable as condition should return false
	result, err := ev.IsConditionTrue("x", map[string]any{"x": nil})
	// expr-lang evaluates nil as-is, not as bool; this should be false
	require.NoError(t, err)
	assert.False(t, result)
}

func TestExpr_SliceAccess(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("list[0]", newTestEvalEnv())
	require.NoError(t, err)
	assert.Equal(t, 1, result)
}

func TestExpr_ConcurrencySafe(t *testing.T) {
	ev := NewExpressionEvaluator()
	var wg sync.WaitGroup
	for i := 0; i < 100; i++ {
		wg.Add(1)
		go func(n int) {
			defer wg.Done()
			env := map[string]any{"n": n}
			result, err := ev.Evaluate("n * 2", env)
			assert.NoError(t, err)
			assert.Equal(t, n*2, result)
		}(i)
	}
	wg.Wait()
}

func TestExpr_EmptyExpression(t *testing.T) {
	ev := NewExpressionEvaluator()
	result, err := ev.Evaluate("", map[string]any{})
	require.NoError(t, err)
	assert.Nil(t, result)
}

// --- ParseExpressions Tests ---

func TestParseExpressions_Single(t *testing.T) {
	segs := ParseExpressions("${e.Name}", "${", "}")
	require.Len(t, segs, 1)
	assert.True(t, segs[0].IsExpression)
	assert.Equal(t, "e.Name", segs[0].Text)
}

func TestParseExpressions_Multiple(t *testing.T) {
	segs := ParseExpressions("${e.First} ${e.Last}", "${", "}")
	require.Len(t, segs, 3)
	assert.True(t, segs[0].IsExpression)
	assert.Equal(t, "e.First", segs[0].Text)
	assert.False(t, segs[1].IsExpression)
	assert.Equal(t, " ", segs[1].Text)
	assert.True(t, segs[2].IsExpression)
	assert.Equal(t, "e.Last", segs[2].Text)
}

func TestParseExpressions_NoExpr(t *testing.T) {
	segs := ParseExpressions("Hello World", "${", "}")
	require.Len(t, segs, 1)
	assert.False(t, segs[0].IsExpression)
	assert.Equal(t, "Hello World", segs[0].Text)
}

func TestParseExpressions_MixedContent(t *testing.T) {
	segs := ParseExpressions("Name: ${e.Name}, Age: ${e.Age}", "${", "}")
	require.Len(t, segs, 4)
	assert.Equal(t, "Name: ", segs[0].Text)
	assert.Equal(t, "e.Name", segs[1].Text)
	assert.Equal(t, ", Age: ", segs[2].Text)
	assert.Equal(t, "e.Age", segs[3].Text)
}

func TestParseExpressions_CustomNotation(t *testing.T) {
	segs := ParseExpressions("{{e.Name}}", "{{", "}}")
	require.Len(t, segs, 1)
	assert.True(t, segs[0].IsExpression)
	assert.Equal(t, "e.Name", segs[0].Text)
}

func TestParseExpressions_Empty(t *testing.T) {
	segs := ParseExpressions("", "${", "}")
	assert.Empty(t, segs)
}

func TestParseExpressions_DefaultNotation(t *testing.T) {
	segs := ParseExpressions("${x}", "", "") // empty begin/end → use defaults
	require.Len(t, segs, 1)
	assert.True(t, segs[0].IsExpression)
	assert.Equal(t, "x", segs[0].Text)
}

// --- IsExpressionOnly / ExtractSingleExpression Tests ---

func TestIsExpressionOnly_True(t *testing.T) {
	assert.True(t, IsExpressionOnly("${e.Name}", "${", "}"))
}

func TestIsExpressionOnly_False_MixedContent(t *testing.T) {
	assert.False(t, IsExpressionOnly("Name: ${e.Name}", "${", "}"))
}

func TestIsExpressionOnly_False_MultipleExpressions(t *testing.T) {
	assert.False(t, IsExpressionOnly("${e.First}${e.Last}", "${", "}"))
}

func TestIsExpressionOnly_False_NoExpression(t *testing.T) {
	assert.False(t, IsExpressionOnly("Hello", "${", "}"))
}

func TestExtractSingleExpression_Success(t *testing.T) {
	expr, ok := ExtractSingleExpression("${e.Payment}", "${", "}")
	assert.True(t, ok)
	assert.Equal(t, "e.Payment", expr)
}

func TestExtractSingleExpression_WithWhitespace(t *testing.T) {
	expr, ok := ExtractSingleExpression("  ${e.Name}  ", "${", "}")
	assert.True(t, ok)
	assert.Equal(t, "e.Name", expr)
}

func TestExtractSingleExpression_NotSingle(t *testing.T) {
	_, ok := ExtractSingleExpression("${a}${b}", "${", "}")
	assert.False(t, ok)
}

func TestExtractSingleExpression_NoExpression(t *testing.T) {
	_, ok := ExtractSingleExpression("Hello", "${", "}")
	assert.False(t, ok)
}
