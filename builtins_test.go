package xlfill

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// --- String functions ---

func TestUpper(t *testing.T) {
	assert.Equal(t, "HELLO", Upper("hello"))
	assert.Equal(t, "HELLO WORLD", Upper("Hello World"))
	assert.Equal(t, "", Upper(""))
}

func TestLower(t *testing.T) {
	assert.Equal(t, "hello", Lower("HELLO"))
	assert.Equal(t, "hello world", Lower("Hello World"))
	assert.Equal(t, "", Lower(""))
}

func TestTitle(t *testing.T) {
	assert.Equal(t, "Hello World", Title("hello world"))
	assert.Equal(t, "Hello", Title("hello"))
	assert.Equal(t, "", Title(""))
}

func TestJoin(t *testing.T) {
	assert.Equal(t, "a, b, c", Join([]string{"a", "b", "c"}, ", "))
	assert.Equal(t, "1-2-3", Join([]any{1, 2, 3}, "-"))
	assert.Equal(t, "hello", Join([]string{"hello"}, ", "))
	assert.Equal(t, "", Join(nil, ", "))
	assert.Equal(t, "", Join([]string{}, ", "))
	// Non-slice returns string representation
	assert.Equal(t, "42", Join(42, ", "))
}

func TestJoin_IntSlice(t *testing.T) {
	assert.Equal(t, "1, 2, 3", Join([]int{1, 2, 3}, ", "))
}

// --- Numeric formatting ---

func TestFormatNumber(t *testing.T) {
	tests := []struct {
		name     string
		val      any
		decimals int
		expected string
	}{
		{"integer", 1234, 0, "1,234"},
		{"float", 1234.567, 2, "1,234.57"},
		{"small number", 42, 0, "42"},
		{"negative", -1234567.89, 2, "-1,234,567.89"},
		{"zero decimals", 1000000, 0, "1,000,000"},
		{"three decimals", 1234.5, 3, "1,234.500"},
		{"float64", 9999.99, 1, "10,000.0"},
		{"zero", 0, 2, "0.00"},
		{"int64", int64(123456789), 0, "123,456,789"},
		{"non-numeric", "hello", 0, "hello"},
		{"nil", nil, 0, "<nil>"},
		{"small float", 0.5, 2, "0.50"},
		{"hundred", 100, 0, "100"},
		{"thousand", 1000, 0, "1,000"},
		{"uint", uint(5000), 0, "5,000"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			assert.Equal(t, tt.expected, FormatNumber(tt.val, tt.decimals))
		})
	}
}

// --- Date formatting ---

func TestFormatDate(t *testing.T) {
	now := time.Date(2024, 3, 15, 14, 30, 0, 0, time.UTC)

	t.Run("time.Time", func(t *testing.T) {
		result := FormatDate(now, "2006-01-02")
		assert.Equal(t, "2024-03-15", result)
	})

	t.Run("time.Time pointer", func(t *testing.T) {
		result := FormatDate(&now, "Jan 2, 2006")
		assert.Equal(t, "Mar 15, 2024", result)
	})

	t.Run("nil pointer", func(t *testing.T) {
		var tp *time.Time
		result := FormatDate(tp, "2006-01-02")
		assert.Equal(t, "", result)
	})

	t.Run("string RFC3339", func(t *testing.T) {
		result := FormatDate("2024-03-15T14:30:00Z", "2006-01-02")
		assert.Equal(t, "2024-03-15", result)
	})

	t.Run("string date only", func(t *testing.T) {
		result := FormatDate("2024-03-15", "Jan 2, 2006")
		assert.Equal(t, "Mar 15, 2024", result)
	})

	t.Run("string datetime", func(t *testing.T) {
		result := FormatDate("2024-03-15 14:30:00", "15:04")
		assert.Equal(t, "14:30", result)
	})

	t.Run("unparseable string", func(t *testing.T) {
		result := FormatDate("not-a-date", "2006-01-02")
		assert.Equal(t, "not-a-date", result)
	})

	t.Run("non-time type", func(t *testing.T) {
		result := FormatDate(42, "2006-01-02")
		assert.Equal(t, "42", result)
	})
}

// --- Coalesce ---

func TestCoalesce(t *testing.T) {
	assert.Equal(t, "hello", Coalesce("hello"))
	assert.Equal(t, "b", Coalesce(nil, "", "b", "c"))
	assert.Equal(t, 42, Coalesce(nil, nil, 42))
	assert.Nil(t, Coalesce(nil, nil))
	assert.Nil(t, Coalesce())
	// Zero int is NOT treated as empty by Coalesce
	assert.Equal(t, 0, Coalesce(0, "default"))
	assert.Equal(t, "first", Coalesce("first", "second"))
}

// --- IfEmpty ---

func TestIfEmpty(t *testing.T) {
	assert.Equal(t, "hello", IfEmpty("hello", "default"))
	assert.Equal(t, "default", IfEmpty(nil, "default"))
	assert.Equal(t, "default", IfEmpty("", "default"))
	assert.Equal(t, "default", IfEmpty(0, "default"))
	assert.Equal(t, "default", IfEmpty(0.0, "default"))
	assert.Equal(t, 42, IfEmpty(42, "default"))
	assert.Equal(t, "value", IfEmpty("value", "default"))
	// Slice is not "empty"
	assert.Equal(t, []int{}, IfEmpty([]int{}, "default"))
}

// --- Aggregate functions ---

func TestSumBy(t *testing.T) {
	items := []map[string]any{
		{"name": "A", "amount": 10},
		{"name": "B", "amount": 20},
		{"name": "C", "amount": 30},
	}
	assert.Equal(t, 60.0, SumBy(items, "amount"))
}

func TestSumBy_Structs(t *testing.T) {
	type Item struct {
		Name   string
		Amount float64
	}
	items := []Item{
		{Name: "A", Amount: 10.5},
		{Name: "B", Amount: 20.3},
	}
	assert.InDelta(t, 30.8, SumBy(items, "Amount"), 0.001)
}

func TestSumBy_Empty(t *testing.T) {
	assert.Equal(t, 0.0, SumBy([]map[string]any{}, "amount"))
	assert.Equal(t, 0.0, SumBy(nil, "amount"))
}

func TestAvgBy(t *testing.T) {
	items := []map[string]any{
		{"val": 10},
		{"val": 20},
		{"val": 30},
	}
	assert.InDelta(t, 20.0, AvgBy(items, "val"), 0.001)
}

func TestAvgBy_Empty(t *testing.T) {
	assert.Equal(t, 0.0, AvgBy([]map[string]any{}, "val"))
	assert.Equal(t, 0.0, AvgBy(nil, "val"))
}

func TestCountBy(t *testing.T) {
	items := []map[string]any{
		{"status": "active"},
		{"status": "inactive"},
		{"status": "active"},
		{"status": "active"},
	}
	assert.Equal(t, 3, CountBy(items, "status", "active"))
	assert.Equal(t, 1, CountBy(items, "status", "inactive"))
	assert.Equal(t, 0, CountBy(items, "status", "unknown"))
}

func TestCountBy_NumericValues(t *testing.T) {
	items := []map[string]any{
		{"score": 100},
		{"score": 200},
		{"score": 100},
	}
	assert.Equal(t, 2, CountBy(items, "score", 100))
}

func TestMinBy(t *testing.T) {
	items := []map[string]any{
		{"val": 30},
		{"val": 10},
		{"val": 20},
	}
	assert.Equal(t, 10, MinBy(items, "val"))
}

func TestMinBy_Strings(t *testing.T) {
	items := []map[string]any{
		{"name": "Charlie"},
		{"name": "Alice"},
		{"name": "Bob"},
	}
	assert.Equal(t, "Alice", MinBy(items, "name"))
}

func TestMinBy_Empty(t *testing.T) {
	assert.Nil(t, MinBy([]map[string]any{}, "val"))
	assert.Nil(t, MinBy(nil, "val"))
}

func TestMaxBy(t *testing.T) {
	items := []map[string]any{
		{"val": 30},
		{"val": 10},
		{"val": 20},
	}
	assert.Equal(t, 30, MaxBy(items, "val"))
}

func TestMaxBy_Strings(t *testing.T) {
	items := []map[string]any{
		{"name": "Charlie"},
		{"name": "Alice"},
		{"name": "Bob"},
	}
	assert.Equal(t, "Charlie", MaxBy(items, "name"))
}

func TestMaxBy_Empty(t *testing.T) {
	assert.Nil(t, MaxBy([]map[string]any{}, "val"))
	assert.Nil(t, MaxBy(nil, "val"))
}

func TestMaxBy_SingleItem(t *testing.T) {
	items := []map[string]any{{"val": 42}}
	assert.Equal(t, 42, MaxBy(items, "val"))
}

// --- i18n ---

func TestMakeI18nFunc(t *testing.T) {
	bundle := map[string]string{
		"greeting": "Hola",
		"farewell": "Adios",
	}
	fn := makeI18nFunc(bundle)

	assert.Equal(t, "Hola", fn("greeting"))
	assert.Equal(t, "Adios", fn("farewell"))
	// Missing key returns key itself
	assert.Equal(t, "unknown.key", fn("unknown.key"))
}

func TestMakeI18nFunc_NilBundle(t *testing.T) {
	fn := makeI18nFunc(nil)
	assert.Equal(t, "any.key", fn("any.key"))
}

// --- Comment ---

func TestComment(t *testing.T) {
	cv := Comment("Hello")
	assert.Equal(t, "Hello", cv.Text)
	assert.Equal(t, "", cv.Author)
}

func TestComment_WithAuthor(t *testing.T) {
	cv := Comment("Note text", "John Doe")
	assert.Equal(t, "Note text", cv.Text)
	assert.Equal(t, "John Doe", cv.Author)
}

func TestCommentValue_Struct(t *testing.T) {
	cv := CommentValue{Text: "test", Author: "author"}
	assert.Equal(t, "test", cv.Text)
	assert.Equal(t, "author", cv.Author)
}

// --- registerBuiltins ---

func TestRegisterBuiltins_AllRegistered(t *testing.T) {
	m := make(map[string]any)
	registerBuiltins(m, nil)

	expectedFuncs := []string{
		"upper", "lower", "title", "join",
		"formatNumber", "formatDate",
		"coalesce", "ifEmpty",
		"sumBy", "avgBy", "countBy", "minBy", "maxBy",
		"t",
	}
	for _, name := range expectedFuncs {
		assert.NotNil(t, m[name], "expected built-in %q to be registered", name)
	}
}

func TestRegisterBuiltins_DoNotOverrideExisting(t *testing.T) {
	m := map[string]any{
		"upper": "custom-value",
	}
	registerBuiltins(m, nil)

	// Should NOT have been overridden
	assert.Equal(t, "custom-value", m["upper"])
}

func TestRegisterBuiltins_WithI18nBundle(t *testing.T) {
	bundle := map[string]string{"hello": "Bonjour"}
	m := make(map[string]any)
	registerBuiltins(m, bundle)

	fn, ok := m["t"].(func(string) string)
	require.True(t, ok)
	assert.Equal(t, "Bonjour", fn("hello"))
}

// --- restoreBuiltin ---

func TestRestoreBuiltin(t *testing.T) {
	m := make(map[string]any)

	assert.True(t, restoreBuiltin(m, "upper", nil))
	assert.NotNil(t, m["upper"])

	assert.True(t, restoreBuiltin(m, "lower", nil))
	assert.True(t, restoreBuiltin(m, "title", nil))
	assert.True(t, restoreBuiltin(m, "join", nil))
	assert.True(t, restoreBuiltin(m, "formatNumber", nil))
	assert.True(t, restoreBuiltin(m, "formatDate", nil))
	assert.True(t, restoreBuiltin(m, "coalesce", nil))
	assert.True(t, restoreBuiltin(m, "ifEmpty", nil))
	assert.True(t, restoreBuiltin(m, "sumBy", nil))
	assert.True(t, restoreBuiltin(m, "avgBy", nil))
	assert.True(t, restoreBuiltin(m, "countBy", nil))
	assert.True(t, restoreBuiltin(m, "minBy", nil))
	assert.True(t, restoreBuiltin(m, "maxBy", nil))
	assert.True(t, restoreBuiltin(m, "t", nil))

	assert.False(t, restoreBuiltin(m, "unknownFunc", nil))
}

// --- Context integration ---

func TestContext_ToMap_IncludesBuiltins(t *testing.T) {
	ctx := NewContext(map[string]any{"name": "test"})
	m := ctx.ToMap()

	assert.NotNil(t, m["hyperlink"])
	assert.NotNil(t, m["comment"])
	assert.NotNil(t, m["upper"])
	assert.NotNil(t, m["lower"])
	assert.NotNil(t, m["title"])
	assert.NotNil(t, m["formatNumber"])
	assert.NotNil(t, m["formatDate"])
	assert.NotNil(t, m["coalesce"])
	assert.NotNil(t, m["ifEmpty"])
	assert.NotNil(t, m["join"])
	assert.NotNil(t, m["sumBy"])
	assert.NotNil(t, m["avgBy"])
	assert.NotNil(t, m["countBy"])
	assert.NotNil(t, m["minBy"])
	assert.NotNil(t, m["maxBy"])
	assert.NotNil(t, m["t"])
}

func TestContext_ToMap_UserDataOverridesBuiltins(t *testing.T) {
	ctx := NewContext(map[string]any{"upper": "my-upper"})
	m := ctx.ToMap()
	assert.Equal(t, "my-upper", m["upper"])
}

func TestContext_CustomFunctions(t *testing.T) {
	myFunc := func(x int) int { return x * 2 }
	ctx := NewContext(map[string]any{"name": "test"}, WithCustomFunctions(map[string]any{
		"double": myFunc,
	}))
	m := ctx.ToMap()
	assert.NotNil(t, m["double"])
}

func TestContext_CustomFunctions_UserDataTakesPrecedence(t *testing.T) {
	ctx := NewContext(
		map[string]any{"myFn": "user-value"},
		WithCustomFunctions(map[string]any{"myFn": func() string { return "custom" }}),
	)
	m := ctx.ToMap()
	assert.Equal(t, "user-value", m["myFn"])
}

func TestContext_I18nBundle(t *testing.T) {
	ctx := NewContext(
		map[string]any{},
		WithI18nBundle(map[string]string{"hello": "Hola"}),
	)
	m := ctx.ToMap()
	fn, ok := m["t"].(func(string) string)
	require.True(t, ok)
	assert.Equal(t, "Hola", fn("hello"))
}

func TestContext_Clone_PreservesCustomFunctions(t *testing.T) {
	ctx := NewContext(
		map[string]any{},
		WithCustomFunctions(map[string]any{"myFn": func() string { return "ok" }}),
	)
	clone := ctx.Clone()
	m := clone.ToMap()
	assert.NotNil(t, m["myFn"])
}

func TestContext_Clone_PreservesI18nBundle(t *testing.T) {
	ctx := NewContext(
		map[string]any{},
		WithI18nBundle(map[string]string{"key": "value"}),
	)
	clone := ctx.Clone()
	m := clone.ToMap()
	fn, ok := m["t"].(func(string) string)
	require.True(t, ok)
	assert.Equal(t, "value", fn("key"))
}

// --- iterateField edge cases ---

func TestIterateField_NonSlice(t *testing.T) {
	called := false
	iterateField("not-a-slice", "field", func(val any) {
		called = true
	})
	assert.False(t, called)
}

func TestIterateField_MissingField(t *testing.T) {
	items := []map[string]any{
		{"name": "A"},
		{"name": "B"},
	}
	var sum float64
	iterateField(items, "amount", func(val any) {
		sum += toFloat64Val(val)
	})
	assert.Equal(t, 0.0, sum)
}

func TestGetFieldValue_StructPointer(t *testing.T) {
	type Item struct {
		Name string
	}
	item := &Item{Name: "test"}
	assert.Equal(t, "test", getFieldValue(item, "Name"))
}

func TestGetFieldValue_NilPointer(t *testing.T) {
	type Item struct{ Name string }
	var item *Item
	assert.Nil(t, getFieldValue(item, "Name"))
}

func TestGetFieldValue_Nil(t *testing.T) {
	assert.Nil(t, getFieldValue(nil, "Name"))
}

func TestGetFieldValue_MapStringAny(t *testing.T) {
	m := map[string]any{"key": "value"}
	assert.Equal(t, "value", getFieldValue(m, "key"))
	assert.Nil(t, getFieldValue(m, "missing"))
}

// --- toFloat64Val edge cases ---

func TestToFloat64Val(t *testing.T) {
	tests := []struct {
		name string
		val  any
		want float64
		nan  bool
	}{
		{"int", 42, 42.0, false},
		{"float64", 3.14, 3.14, false},
		{"int64", int64(100), 100.0, false},
		{"uint", uint(5), 5.0, false},
		{"string", "hello", 0, true},
		{"nil", nil, 0, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := toFloat64Val(tt.val)
			if tt.nan {
				assert.True(t, result != result, "expected NaN") // NaN != NaN
			} else {
				assert.InDelta(t, tt.want, result, 0.001)
			}
		})
	}
}

// --- WithFunction option ---

func TestWithFunction_Option(t *testing.T) {
	opts := defaultOptions()
	WithFunction("myFn", func() string { return "hello" })(opts)

	assert.NotNil(t, opts.customFunctions)
	assert.NotNil(t, opts.customFunctions["myFn"])
}

func TestWithFunction_MultipleOptions(t *testing.T) {
	opts := defaultOptions()
	WithFunction("fn1", func() {})(opts)
	WithFunction("fn2", func() {})(opts)

	assert.Len(t, opts.customFunctions, 2)
}

// --- WithI18n option ---

func TestWithI18n_Option(t *testing.T) {
	opts := defaultOptions()
	bundle := map[string]string{"key": "val"}
	WithI18n(bundle)(opts)

	assert.Equal(t, bundle, opts.i18nBundle)
}

// --- FormatNumber edge cases ---

func TestFormatNumber_NegativeSmall(t *testing.T) {
	assert.Equal(t, "-42", FormatNumber(-42, 0))
}

func TestFormatNumber_LargeNumber(t *testing.T) {
	assert.Equal(t, "1,234,567,890", FormatNumber(1234567890, 0))
}

// --- SumBy with float fields ---

func TestSumBy_FloatFields(t *testing.T) {
	items := []map[string]any{
		{"price": 10.5},
		{"price": 20.3},
		{"price": 30.2},
	}
	assert.InDelta(t, 61.0, SumBy(items, "price"), 0.001)
}

// --- AvgBy with single item ---

func TestAvgBy_SingleItem(t *testing.T) {
	items := []map[string]any{{"val": 42}}
	assert.Equal(t, 42.0, AvgBy(items, "val"))
}

// --- CountBy empty ---

func TestCountBy_Empty(t *testing.T) {
	assert.Equal(t, 0, CountBy(nil, "field", "value"))
	assert.Equal(t, 0, CountBy([]map[string]any{}, "field", "value"))
}
