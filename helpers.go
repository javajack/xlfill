package xlfill

import (
	"encoding/json"
	"fmt"
	"io"
	"reflect"
)

// StructSliceToData converts a slice of structs to []map[string]any for use with xlfill.
// Field names become map keys. Only exported fields are included.
// Pointer-to-struct elements are dereferenced automatically.
func StructSliceToData[T any](items []T) []map[string]any {
	result := make([]map[string]any, len(items))
	for i, item := range items {
		result[i] = structToMap(item)
	}
	return result
}

func structToMap(v any) map[string]any {
	val := reflect.ValueOf(v)
	if val.Kind() == reflect.Ptr {
		val = val.Elem()
	}
	if val.Kind() != reflect.Struct {
		return map[string]any{"value": v}
	}

	t := val.Type()
	m := make(map[string]any, t.NumField())
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		if !f.IsExported() {
			continue
		}
		m[f.Name] = val.Field(i).Interface()
	}
	return m
}

// JSONToData parses JSON from a reader into a map[string]any for use with xlfill.
func JSONToData(r io.Reader) (map[string]any, error) {
	var result map[string]any
	if err := json.NewDecoder(r).Decode(&result); err != nil {
		return nil, fmt.Errorf("json decode: %w", err)
	}
	return result, nil
}

// RowScanner is an interface compatible with *sql.Rows for reading database results.
type RowScanner interface {
	Columns() ([]string, error)
	Next() bool
	Scan(dest ...any) error
}

// SQLRowsToData converts database rows to []map[string]any for use with xlfill.
// Each row becomes a map keyed by column name.
func SQLRowsToData(rows RowScanner) ([]map[string]any, error) {
	cols, err := rows.Columns()
	if err != nil {
		return nil, fmt.Errorf("get columns: %w", err)
	}

	var result []map[string]any
	for rows.Next() {
		values := make([]any, len(cols))
		ptrs := make([]any, len(cols))
		for i := range values {
			ptrs[i] = &values[i]
		}

		if err := rows.Scan(ptrs...); err != nil {
			return nil, fmt.Errorf("scan row: %w", err)
		}

		row := make(map[string]any, len(cols))
		for i, col := range cols {
			row[col] = values[i]
		}
		result = append(result, row)
	}
	return result, nil
}
