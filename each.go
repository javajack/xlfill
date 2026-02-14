package xlfill

import (
	"fmt"
	"reflect"
	"strings"
)

// EachCommand implements the jx:each command for iterating over collections.
type EachCommand struct {
	Items     string // expression for collection (e.g., "employees")
	Var       string // loop variable name (e.g., "e")
	VarIndex  string // optional index variable name (e.g., "idx")
	Direction string // "DOWN" (default) or "RIGHT"
	Area      *Area  // the template area to repeat for each item

	// Advanced (Phase 10)
	Select     string // filter expression
	GroupBy    string // grouping property
	GroupOrder string // "ASC" or "DESC"
	OrderBy    string // sort specification
	MultiSheet string // sheet names variable
}

func (c *EachCommand) Name() string { return "each" }
func (c *EachCommand) Reset()       {}

// newEachCommandFromAttrs creates an EachCommand from parsed attributes.
func newEachCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &EachCommand{
		Items:      attrs["items"],
		Var:        attrs["var"],
		VarIndex:   attrs["varIndex"],
		Direction:  strings.ToUpper(attrs["direction"]),
		Select:     attrs["select"],
		GroupBy:    attrs["groupBy"],
		GroupOrder: attrs["groupOrder"],
		OrderBy:    attrs["orderBy"],
		MultiSheet: attrs["multisheet"],
	}
	if cmd.Items == "" {
		return nil, fmt.Errorf("each command requires 'items' attribute")
	}
	if cmd.Var == "" {
		return nil, fmt.Errorf("each command requires 'var' attribute")
	}
	if cmd.Direction == "" {
		cmd.Direction = "DOWN"
	}
	return cmd, nil
}

// ApplyAt executes the each command at the given target cell.
func (c *EachCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	// Evaluate items expression
	itemsVal, err := ctx.Evaluate(c.Items)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate items %q: %w", c.Items, err)
	}

	// Convert to iterable slice
	items, err := toSlice(itemsVal)
	if err != nil {
		return ZeroSize, fmt.Errorf("items %q is not iterable: %w", c.Items, err)
	}

	if len(items) == 0 {
		return ZeroSize, nil
	}

	// Apply select filter
	if c.Select != "" {
		items, err = c.filterItems(items, ctx)
		if err != nil {
			return ZeroSize, err
		}
		if len(items) == 0 {
			return ZeroSize, nil
		}
	}

	// Apply groupBy — transforms items into []GroupData
	if c.GroupBy != "" {
		items = c.groupItems(items)
	}

	// Apply orderBy
	if c.OrderBy != "" {
		items, err = c.sortItems(items)
		if err != nil {
			return ZeroSize, err
		}
	}

	if c.Area == nil {
		return ZeroSize, fmt.Errorf("each command has no area")
	}

	// Multisheet mode: each item gets its own sheet
	if c.MultiSheet != "" {
		return c.applyMultiSheet(cellRef, ctx, transformer, items)
	}

	// Iterate
	isRight := c.Direction == "RIGHT"
	totalSize := ZeroSize

	for i, item := range items {
		// Set loop variable
		var rv *RunVar
		if c.VarIndex != "" {
			rv = NewRunVarWithIndex(ctx, c.Var, c.VarIndex)
			rv.SetWithIndex(item, i)
		} else {
			rv = NewRunVar(ctx, c.Var)
			rv.Set(item)
		}

		// Calculate target cell for this iteration
		var iterTarget CellRef
		if isRight {
			iterTarget = NewCellRef(cellRef.Sheet, cellRef.Row, cellRef.Col+totalSize.Width)
		} else {
			iterTarget = NewCellRef(cellRef.Sheet, cellRef.Row+totalSize.Height, cellRef.Col)
		}

		// Apply area at target
		iterSize, err := c.Area.ApplyAt(iterTarget, ctx)
		rv.Close()
		if err != nil {
			return ZeroSize, fmt.Errorf("each iteration %d: %w", i, err)
		}

		// Accumulate size
		if isRight {
			totalSize.Width += iterSize.Width
			if iterSize.Height > totalSize.Height {
				totalSize.Height = iterSize.Height
			}
		} else {
			totalSize.Height += iterSize.Height
			if iterSize.Width > totalSize.Width {
				totalSize.Width = iterSize.Width
			}
		}
	}

	return totalSize, nil
}

// applyMultiSheet processes each item on a separate sheet.
// The multisheet attribute holds the name of a context variable containing sheet names.
func (c *EachCommand) applyMultiSheet(cellRef CellRef, ctx *Context, transformer Transformer, items []any) (Size, error) {
	// Evaluate multisheet to get sheet names
	sheetNamesVal, err := ctx.Evaluate(c.MultiSheet)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate multisheet %q: %w", c.MultiSheet, err)
	}
	sheetNames, err := toStringSlice(sheetNamesVal)
	if err != nil {
		return ZeroSize, fmt.Errorf("multisheet %q must be a string slice: %w", c.MultiSheet, err)
	}

	templateSheet := cellRef.Sheet
	lastSize := ZeroSize

	for i, item := range items {
		// Determine sheet name
		var sheetName string
		if i < len(sheetNames) {
			sheetName = sheetNames[i]
		} else {
			sheetName = fmt.Sprintf("%s_%d", templateSheet, i+1)
		}
		sheetName = SafeSheetName(sheetName)

		// Copy template sheet
		if err := transformer.CopySheet(templateSheet, sheetName); err != nil {
			return ZeroSize, fmt.Errorf("copy sheet for multisheet item %d: %w", i, err)
		}

		// Set loop variable
		var rv *RunVar
		if c.VarIndex != "" {
			rv = NewRunVarWithIndex(ctx, c.Var, c.VarIndex)
			rv.SetWithIndex(item, i)
		} else {
			rv = NewRunVar(ctx, c.Var)
			rv.Set(item)
		}

		// Create a target on the new sheet at the same position
		target := NewCellRef(sheetName, cellRef.Row, cellRef.Col)

		// Process the area on the new sheet — we need to read cell data from the new sheet.
		// Since the sheet was copied, the transformer already has the data.
		// We use the template area's size but target the new sheet.
		iterSize, err := c.Area.ApplyAt(target, ctx)
		rv.Close()
		if err != nil {
			return ZeroSize, fmt.Errorf("multisheet iteration %d (sheet %s): %w", i, sheetName, err)
		}
		lastSize = iterSize
	}

	// Delete the template sheet (it was the source for copies)
	transformer.DeleteSheet(templateSheet)

	return lastSize, nil
}

// toStringSlice converts a value to []string.
func toStringSlice(val any) ([]string, error) {
	if val == nil {
		return nil, nil
	}
	items, err := toSlice(val)
	if err != nil {
		return nil, err
	}
	result := make([]string, len(items))
	for i, item := range items {
		result[i] = fmt.Sprintf("%v", item)
	}
	return result, nil
}

// filterItems applies the select expression to filter items.
func (c *EachCommand) filterItems(items []any, ctx *Context) ([]any, error) {
	var filtered []any
	for _, item := range items {
		rv := NewRunVar(ctx, c.Var)
		rv.Set(item)
		ok, err := ctx.IsConditionTrue(c.Select)
		rv.Close()
		if err != nil {
			return nil, fmt.Errorf("select filter %q: %w", c.Select, err)
		}
		if ok {
			filtered = append(filtered, item)
		}
	}
	return filtered, nil
}

// sortItems sorts items by the orderBy specification.
func (c *EachCommand) sortItems(items []any) ([]any, error) {
	// Parse orderBy: "e.Name ASC, e.Payment DESC"
	specs := parseOrderBy(c.OrderBy, c.Var)
	if len(specs) == 0 {
		return items, nil
	}
	sortByFields(items, specs)
	return items, nil
}

// GroupData represents a group of items sharing a common key value.
// Used with groupBy: ${g.item.Department} accesses the key, ${g.items} iterates group members.
type GroupData struct {
	Item  any   // the first item in the group (or representative)
	Items []any // all items in this group
}

// groupItems groups items by the groupBy property and returns []GroupData wrapped as []any.
func (c *EachCommand) groupItems(items []any) []any {
	field := c.GroupBy
	// Strip var prefix (e.g., "e.Department" → "Department")
	prefix := c.Var + "."
	if strings.HasPrefix(field, prefix) {
		field = field[len(prefix):]
	}

	// Maintain insertion order
	type groupEntry struct {
		key   any
		items []any
	}
	var groups []groupEntry
	keyIndex := map[string]int{} // string representation → index

	for _, item := range items {
		val := getField(item, field)
		keyStr := fmt.Sprintf("%v", val)
		if idx, ok := keyIndex[keyStr]; ok {
			groups[idx].items = append(groups[idx].items, item)
		} else {
			keyIndex[keyStr] = len(groups)
			groups = append(groups, groupEntry{key: val, items: []any{item}})
		}
	}

	// Sort groups if groupOrder specified
	if c.GroupOrder != "" {
		orderDesc := strings.Contains(strings.ToUpper(c.GroupOrder), "DESC")
		ignoreCase := strings.Contains(strings.ToUpper(c.GroupOrder), "IGNORECASE") ||
			strings.Contains(strings.ToUpper(c.GroupOrder), "IGNORE_CASE")

		// Stable insertion sort
		for i := 1; i < len(groups); i++ {
			key := groups[i]
			j := i - 1
			for j >= 0 && compareGroupKeys(groups[j].key, key.key, orderDesc, ignoreCase) > 0 {
				groups[j+1] = groups[j]
				j--
			}
			groups[j+1] = key
		}
	}

	// Convert to []any of GroupData
	result := make([]any, len(groups))
	for i, g := range groups {
		result[i] = GroupData{Item: g.items[0], Items: g.items}
	}
	return result
}

// compareGroupKeys compares two group keys for sorting.
func compareGroupKeys(a, b any, desc, ignoreCase bool) int {
	var cmp int
	if ignoreCase {
		sa := strings.ToLower(fmt.Sprintf("%v", a))
		sb := strings.ToLower(fmt.Sprintf("%v", b))
		if sa < sb {
			cmp = -1
		} else if sa > sb {
			cmp = 1
		}
	} else {
		cmp = compareValues(a, b)
	}
	if desc {
		cmp = -cmp
	}
	return cmp
}

// orderBySpec represents a single sort field with direction.
type orderBySpec struct {
	field string // field name without var prefix (e.g., "Name")
	desc  bool   // true for DESC
}

// parseOrderBy parses an orderBy string like "e.Name ASC, e.Payment DESC".
func parseOrderBy(spec string, varName string) []orderBySpec {
	if strings.TrimSpace(spec) == "" {
		return nil
	}
	parts := strings.Split(spec, ",")
	var specs []orderBySpec
	prefix := varName + "."
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p == "" {
			continue
		}
		tokens := strings.Fields(p)
		field := tokens[0]
		// Strip var prefix
		if strings.HasPrefix(field, prefix) {
			field = field[len(prefix):]
		}
		desc := false
		if len(tokens) > 1 && strings.EqualFold(tokens[1], "DESC") {
			desc = true
		}
		specs = append(specs, orderBySpec{field: field, desc: desc})
	}
	return specs
}

// sortByFields sorts items in place by the given field specs.
func sortByFields(items []any, specs []orderBySpec) {
	if len(specs) == 0 || len(items) <= 1 {
		return
	}
	// Simple insertion sort (stable) for template data sizes
	for i := 1; i < len(items); i++ {
		key := items[i]
		j := i - 1
		for j >= 0 && compareBySpecs(items[j], key, specs) > 0 {
			items[j+1] = items[j]
			j--
		}
		items[j+1] = key
	}
}

// compareBySpecs compares two items by the orderBy specs.
func compareBySpecs(a, b any, specs []orderBySpec) int {
	for _, s := range specs {
		va := getField(a, s.field)
		vb := getField(b, s.field)
		cmp := compareValues(va, vb)
		if s.desc {
			cmp = -cmp
		}
		if cmp != 0 {
			return cmp
		}
	}
	return 0
}

// getField extracts a field value from a struct or map by name.
func getField(item any, field string) any {
	if item == nil {
		return nil
	}
	// Try map first
	if m, ok := item.(map[string]any); ok {
		return m[field]
	}
	// Try struct via reflection
	v := reflect.ValueOf(item)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	if v.Kind() == reflect.Struct {
		f := v.FieldByName(field)
		if f.IsValid() {
			return f.Interface()
		}
	}
	return nil
}

// compareValues compares two values for ordering.
func compareValues(a, b any) int {
	if a == nil && b == nil {
		return 0
	}
	if a == nil {
		return -1
	}
	if b == nil {
		return 1
	}
	// Convert to float64 for numeric comparison
	fa, aOk := toFloat64(a)
	fb, bOk := toFloat64(b)
	if aOk && bOk {
		if fa < fb {
			return -1
		}
		if fa > fb {
			return 1
		}
		return 0
	}
	// Fall back to string comparison
	sa := fmt.Sprintf("%v", a)
	sb := fmt.Sprintf("%v", b)
	if sa < sb {
		return -1
	}
	if sa > sb {
		return 1
	}
	return 0
}

// toFloat64 attempts to convert a value to float64.
func toFloat64(v any) (float64, bool) {
	switch n := v.(type) {
	case int:
		return float64(n), true
	case int8:
		return float64(n), true
	case int16:
		return float64(n), true
	case int32:
		return float64(n), true
	case int64:
		return float64(n), true
	case float32:
		return float64(n), true
	case float64:
		return n, true
	}
	return 0, false
}

// toSlice converts any iterable value to a []any slice.
func toSlice(val any) ([]any, error) {
	if val == nil {
		return nil, nil
	}

	v := reflect.ValueOf(val)
	switch v.Kind() {
	case reflect.Slice, reflect.Array:
		result := make([]any, v.Len())
		for i := 0; i < v.Len(); i++ {
			result[i] = v.Index(i).Interface()
		}
		return result, nil
	default:
		return nil, fmt.Errorf("cannot iterate over %T", val)
	}
}
