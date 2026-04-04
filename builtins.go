package xlfill

import (
	"fmt"
	"math"
	"reflect"
	"strings"
	"time"
)

// registerBuiltins registers all built-in template functions into the map.
// Only sets a key if it is not already present (user data takes precedence).
func registerBuiltins(m map[string]any, i18nBundle map[string]string) {
	set := func(name string, fn any) {
		if _, ok := m[name]; !ok {
			m[name] = fn
		}
	}

	// String functions
	set("upper", Upper)
	set("lower", Lower)
	set("title", Title)
	set("join", Join)

	// Numeric formatting
	set("formatNumber", FormatNumber)

	// Date formatting
	set("formatDate", FormatDate)

	// Coalesce / default
	set("coalesce", Coalesce)
	set("ifEmpty", IfEmpty)

	// Aggregate functions
	set("sumBy", SumBy)
	set("avgBy", AvgBy)
	set("countBy", CountBy)
	set("minBy", MinBy)
	set("maxBy", MaxBy)

	// i18n
	if i18nBundle != nil {
		set("t", makeI18nFunc(i18nBundle))
	} else {
		set("t", makeI18nFunc(nil))
	}
}

// restoreBuiltin restores a built-in function key in the cached map during
// differential updates. Returns true if the key was a known built-in.
func restoreBuiltin(m map[string]any, key string, i18nBundle map[string]string) bool {
	switch key {
	case "upper":
		m[key] = Upper
	case "lower":
		m[key] = Lower
	case "title":
		m[key] = Title
	case "join":
		m[key] = Join
	case "formatNumber":
		m[key] = FormatNumber
	case "formatDate":
		m[key] = FormatDate
	case "coalesce":
		m[key] = Coalesce
	case "ifEmpty":
		m[key] = IfEmpty
	case "sumBy":
		m[key] = SumBy
	case "avgBy":
		m[key] = AvgBy
	case "countBy":
		m[key] = CountBy
	case "minBy":
		m[key] = MinBy
	case "maxBy":
		m[key] = MaxBy
	case "t":
		m[key] = makeI18nFunc(i18nBundle)
	default:
		return false
	}
	return true
}

// --- String functions ---

// Upper returns the uppercase version of a string.
func Upper(s string) string { return strings.ToUpper(s) }

// Lower returns the lowercase version of a string.
func Lower(s string) string { return strings.ToLower(s) }

// Title returns the title-case version of a string.
func Title(s string) string { return strings.Title(s) } //nolint:staticcheck

// Join joins slice elements with a separator.
// Items can be a []string, []any, or any slice type.
func Join(items any, sep string) string {
	if items == nil {
		return ""
	}
	v := reflect.ValueOf(items)
	if v.Kind() != reflect.Slice {
		return fmt.Sprintf("%v", items)
	}
	parts := make([]string, v.Len())
	for i := 0; i < v.Len(); i++ {
		parts[i] = fmt.Sprintf("%v", v.Index(i).Interface())
	}
	return strings.Join(parts, sep)
}

// --- Numeric formatting ---

// FormatNumber formats a number with comma separators and the specified number of decimal places.
func FormatNumber(val any, decimals int) string {
	f := toFloat64Val(val)
	if math.IsNaN(f) {
		return fmt.Sprintf("%v", val)
	}

	// Format with decimals
	formatted := fmt.Sprintf("%.*f", decimals, f)

	// Split into integer and decimal parts
	parts := strings.SplitN(formatted, ".", 2)
	intPart := parts[0]

	// Handle negative sign
	negative := false
	if strings.HasPrefix(intPart, "-") {
		negative = true
		intPart = intPart[1:]
	}

	// Add comma separators
	n := len(intPart)
	if n > 3 {
		var b strings.Builder
		remainder := n % 3
		if remainder > 0 {
			b.WriteString(intPart[:remainder])
			if n > remainder {
				b.WriteByte(',')
			}
		}
		for i := remainder; i < n; i += 3 {
			if i > remainder {
				b.WriteByte(',')
			}
			b.WriteString(intPart[i : i+3])
		}
		intPart = b.String()
	}

	if negative {
		intPart = "-" + intPart
	}

	if len(parts) > 1 {
		return intPart + "." + parts[1]
	}
	return intPart
}

// --- Date formatting ---

// FormatDate formats a time.Time or string date using a Go time layout.
// If date is a string, it attempts to parse it using common formats.
func FormatDate(date any, layout string) string {
	var t time.Time
	switch d := date.(type) {
	case time.Time:
		t = d
	case *time.Time:
		if d == nil {
			return ""
		}
		t = *d
	case string:
		var err error
		// Try common formats
		for _, fmt := range []string{
			time.RFC3339,
			"2006-01-02T15:04:05",
			"2006-01-02 15:04:05",
			"2006-01-02",
			"01/02/2006",
			"02/01/2006",
			"Jan 2, 2006",
			"January 2, 2006",
		} {
			t, err = time.Parse(fmt, d)
			if err == nil {
				break
			}
		}
		if err != nil {
			return d // return original string if unparseable
		}
	default:
		return fmt.Sprintf("%v", date)
	}
	return t.Format(layout)
}

// --- Coalesce / default ---

// Coalesce returns the first non-nil, non-empty value from the arguments.
func Coalesce(values ...any) any {
	for _, v := range values {
		if v != nil {
			// Check for empty string
			if s, ok := v.(string); ok && s == "" {
				continue
			}
			return v
		}
	}
	return nil
}

// IfEmpty returns defaultVal if val is nil, empty string, or the zero value of its type.
func IfEmpty(val any, defaultVal any) any {
	if val == nil {
		return defaultVal
	}
	if s, ok := val.(string); ok && s == "" {
		return defaultVal
	}
	// Check numeric zero
	v := reflect.ValueOf(val)
	if v.Kind() >= reflect.Int && v.Kind() <= reflect.Float64 && v.IsZero() {
		return defaultVal
	}
	return val
}

// --- Aggregate functions ---

// SumBy sums a numeric field across a slice of maps or structs.
func SumBy(items any, field string) float64 {
	var sum float64
	iterateField(items, field, func(val any) {
		sum += toFloat64Val(val)
	})
	return sum
}

// AvgBy computes the average of a numeric field across a slice of maps or structs.
func AvgBy(items any, field string) float64 {
	var sum float64
	var count int
	iterateField(items, field, func(val any) {
		sum += toFloat64Val(val)
		count++
	})
	if count == 0 {
		return 0
	}
	return sum / float64(count)
}

// CountBy counts items where the specified field equals the given value.
func CountBy(items any, field string, value any) int {
	valueStr := fmt.Sprintf("%v", value)
	var count int
	iterateField(items, field, func(val any) {
		if fmt.Sprintf("%v", val) == valueStr {
			count++
		}
	})
	return count
}

// MinBy returns the minimum value of a field across a slice of maps or structs.
func MinBy(items any, field string) any {
	var minVal any
	first := true
	iterateField(items, field, func(val any) {
		if first {
			minVal = val
			first = false
			return
		}
		if compareValues(val, minVal) < 0 {
			minVal = val
		}
	})
	return minVal
}

// MaxBy returns the maximum value of a field across a slice of maps or structs.
func MaxBy(items any, field string) any {
	var maxVal any
	first := true
	iterateField(items, field, func(val any) {
		if first {
			maxVal = val
			first = false
			return
		}
		if compareValues(val, maxVal) > 0 {
			maxVal = val
		}
	})
	return maxVal
}

// --- i18n ---

// makeI18nFunc creates a translation function that looks up keys in the bundle.
// If the key is not found, the key itself is returned.
func makeI18nFunc(bundle map[string]string) func(key string) string {
	return func(key string) string {
		if bundle != nil {
			if v, ok := bundle[key]; ok {
				return v
			}
		}
		return key
	}
}

// --- Comment ---

// CommentValue represents a cell value with an associated comment.
// When an expression evaluates to this type, the transformer writes both
// the cell text and an Excel comment.
type CommentValue struct {
	Text   string
	Author string
}

// Comment creates a CommentValue for use in template expressions.
// Usage in template: ${comment("This is a note")} or ${comment("Note", "Author")}
func Comment(text string, author ...string) CommentValue {
	cv := CommentValue{Text: text}
	if len(author) > 0 {
		cv.Author = author[0]
	}
	return cv
}

// --- Helpers ---

// toFloat64Val converts a numeric value to float64, returning NaN for non-numeric types.
// This wraps the existing toFloat64 from each.go which returns (float64, bool).
func toFloat64Val(val any) float64 {
	f, ok := toFloat64(val)
	if !ok {
		return math.NaN()
	}
	return f
}

// iterateField iterates over a slice and calls fn with the value of the named field
// for each element. Supports slices of maps (map[string]any) and structs.
func iterateField(items any, field string, fn func(any)) {
	if items == nil {
		return
	}
	v := reflect.ValueOf(items)
	if v.Kind() != reflect.Slice {
		return
	}
	for i := 0; i < v.Len(); i++ {
		elem := v.Index(i).Interface()
		val := getFieldValue(elem, field)
		if val != nil {
			fn(val)
		}
	}
}

// getFieldValue extracts a named field from a map or struct.
func getFieldValue(item any, field string) any {
	if item == nil {
		return nil
	}

	// Try map[string]any first (most common in template data)
	if m, ok := item.(map[string]any); ok {
		return m[field]
	}

	// Try reflection for struct or other map types
	v := reflect.ValueOf(item)
	if v.Kind() == reflect.Ptr {
		if v.IsNil() {
			return nil
		}
		v = v.Elem()
	}

	switch v.Kind() {
	case reflect.Map:
		key := reflect.ValueOf(field)
		result := v.MapIndex(key)
		if result.IsValid() {
			return result.Interface()
		}
	case reflect.Struct:
		fv := v.FieldByName(field)
		if fv.IsValid() {
			return fv.Interface()
		}
	}
	return nil
}

// Note: compareValues is defined in each.go and shared across the package.
