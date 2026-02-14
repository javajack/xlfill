package goxls

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestContext_PutGetVar(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 10})
	assert.Equal(t, 10, ctx.GetVar("x"))

	ctx.PutVar("y", "hello")
	assert.Equal(t, "hello", ctx.GetVar("y"))
}

func TestContext_RemoveVar(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 10})
	ctx.RemoveVar("x")
	assert.Nil(t, ctx.GetVar("x"))
}

func TestContext_ContainsVar(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 10})
	assert.True(t, ctx.ContainsVar("x"))
	assert.False(t, ctx.ContainsVar("y"))
}

func TestContext_ToMap(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 10})
	ctx.setRunVar("y", 20)
	m := ctx.ToMap()
	assert.Equal(t, 10, m["x"])
	assert.Equal(t, 20, m["y"])
}

func TestContext_ToMap_RunVarOverridesData(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 10})
	ctx.setRunVar("x", 99)
	m := ctx.ToMap()
	assert.Equal(t, 99, m["x"])
}

func TestContext_Evaluate(t *testing.T) {
	ctx := NewContext(map[string]any{
		"e": testEmployee{Name: "Bob", Payment: 3000},
	})
	result, err := ctx.Evaluate("e.Name")
	require.NoError(t, err)
	assert.Equal(t, "Bob", result)
}

func TestContext_IsConditionTrue(t *testing.T) {
	ctx := NewContext(map[string]any{
		"e": testEmployee{Payment: 5000},
	})
	result, err := ctx.IsConditionTrue("e.Payment > 2000")
	require.NoError(t, err)
	assert.True(t, result)
}

func TestContext_NilData(t *testing.T) {
	ctx := NewContext(nil)
	assert.NotNil(t, ctx)
	assert.Nil(t, ctx.GetVar("anything"))
}

func TestContext_CustomNotation(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 42}, WithNotation("{{", "}}"))
	val, ct, err := ctx.EvaluateCellValue("{{x}}")
	require.NoError(t, err)
	assert.Equal(t, 42, val)
	assert.Equal(t, CellNumber, ct)
}

// --- RunVar Tests ---

func TestContext_RunVarScope(t *testing.T) {
	ctx := NewContext(map[string]any{})
	ctx.setRunVar("e", "original")

	rv := NewRunVar(ctx, "e")
	rv.Set("override")
	assert.Equal(t, "override", ctx.GetVar("e"))

	rv.Close()
	assert.Equal(t, "original", ctx.GetVar("e"))
}

func TestContext_RunVarScope_NewVar(t *testing.T) {
	ctx := NewContext(map[string]any{})

	rv := NewRunVar(ctx, "e")
	rv.Set("value")
	assert.Equal(t, "value", ctx.GetVar("e"))

	rv.Close()
	// Variable should be removed since it didn't exist before
	assert.Nil(t, ctx.GetVar("e"))
	assert.False(t, ctx.ContainsVar("e"))
}

func TestContext_RunVarWithIndex(t *testing.T) {
	ctx := NewContext(map[string]any{})

	rv := NewRunVarWithIndex(ctx, "e", "idx")
	rv.SetWithIndex("item1", 0)
	assert.Equal(t, "item1", ctx.GetVar("e"))
	assert.Equal(t, 0, ctx.GetVar("idx"))

	rv.SetWithIndex("item2", 1)
	assert.Equal(t, "item2", ctx.GetVar("e"))
	assert.Equal(t, 1, ctx.GetVar("idx"))

	rv.Close()
	assert.Nil(t, ctx.GetVar("e"))
	assert.Nil(t, ctx.GetVar("idx"))
}

func TestContext_RunVarNested(t *testing.T) {
	ctx := NewContext(map[string]any{})

	// Outer loop sets "dept"
	rvOuter := NewRunVar(ctx, "dept")
	rvOuter.Set("Engineering")

	// Inner loop also uses a var "e"
	rvInner := NewRunVar(ctx, "e")
	rvInner.Set("Alice")
	assert.Equal(t, "Alice", ctx.GetVar("e"))
	assert.Equal(t, "Engineering", ctx.GetVar("dept"))

	rvInner.Close()
	assert.Nil(t, ctx.GetVar("e"))
	assert.Equal(t, "Engineering", ctx.GetVar("dept"))

	rvOuter.Close()
	assert.Nil(t, ctx.GetVar("dept"))
}

func TestContext_RunVarNested_SameVar(t *testing.T) {
	ctx := NewContext(map[string]any{})

	// Outer sets "e" to outer value
	rvOuter := NewRunVar(ctx, "e")
	rvOuter.Set("outer")

	// Inner also sets "e" (nested groupBy scenario)
	rvInner := NewRunVar(ctx, "e")
	rvInner.Set("inner")
	assert.Equal(t, "inner", ctx.GetVar("e"))

	rvInner.Close()
	assert.Equal(t, "outer", ctx.GetVar("e"))

	rvOuter.Close()
	assert.Nil(t, ctx.GetVar("e"))
}

// --- EvaluateCellValue Tests ---

func TestContext_EvaluateCellValue_Expression(t *testing.T) {
	ctx := NewContext(map[string]any{
		"e": testEmployee{Name: "Alice", Payment: 5000},
	})
	val, ct, err := ctx.EvaluateCellValue("${e.Name}")
	require.NoError(t, err)
	assert.Equal(t, "Alice", val)
	assert.Equal(t, CellString, ct)
}

func TestContext_EvaluateCellValue_Number(t *testing.T) {
	ctx := NewContext(map[string]any{
		"e": testEmployee{Payment: 5000},
	})
	val, ct, err := ctx.EvaluateCellValue("${e.Payment}")
	require.NoError(t, err)
	assert.Equal(t, 5000.0, val)
	assert.Equal(t, CellNumber, ct)
}

func TestContext_EvaluateCellValue_Bool(t *testing.T) {
	ctx := NewContext(map[string]any{
		"e": testEmployee{Active: true},
	})
	val, ct, err := ctx.EvaluateCellValue("${e.Active}")
	require.NoError(t, err)
	assert.Equal(t, true, val)
	assert.Equal(t, CellBoolean, ct)
}

func TestContext_EvaluateCellValue_Mixed(t *testing.T) {
	ctx := NewContext(map[string]any{
		"e": testEmployee{Name: "Alice"},
	})
	val, ct, err := ctx.EvaluateCellValue("Name: ${e.Name}")
	require.NoError(t, err)
	assert.Equal(t, "Name: Alice", val)
	assert.Equal(t, CellString, ct) // always string for mixed
}

func TestContext_EvaluateCellValue_NoExpression(t *testing.T) {
	ctx := NewContext(map[string]any{})
	val, ct, err := ctx.EvaluateCellValue("Hello World")
	require.NoError(t, err)
	assert.Equal(t, "Hello World", val)
	assert.Equal(t, CellString, ct)
}

func TestContext_EvaluateCellValue_NilResult(t *testing.T) {
	ctx := NewContext(map[string]any{"x": nil})
	val, ct, err := ctx.EvaluateCellValue("${x}")
	require.NoError(t, err)
	assert.Nil(t, val)
	assert.Equal(t, CellBlank, ct)
}

func TestContext_EvaluateCellValue_RunVarVisible(t *testing.T) {
	ctx := NewContext(map[string]any{})
	rv := NewRunVar(ctx, "e")
	rv.Set(testEmployee{Name: "Bob"})

	val, _, err := ctx.EvaluateCellValue("${e.Name}")
	require.NoError(t, err)
	assert.Equal(t, "Bob", val)

	rv.Close()
}
