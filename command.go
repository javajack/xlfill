package xlfill

import "sort"

// Command represents a template processing command (jx:each, jx:if, etc.).
type Command interface {
	Name() string
	ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error)
	Reset()
}

// CommandFactory creates a Command from parsed attributes.
type CommandFactory func(attrs map[string]string) (Command, error)

// CommandRegistry maps command names to their factories.
type CommandRegistry struct {
	factories map[string]CommandFactory
}

// NewCommandRegistry creates a registry with built-in commands.
func NewCommandRegistry() *CommandRegistry {
	r := &CommandRegistry{
		factories: make(map[string]CommandFactory),
	}
	r.Register("each", newEachCommandFromAttrs)
	r.Register("if", newIfCommandFromAttrs)
	r.Register("grid", newGridCommandFromAttrs)
	r.Register("image", newImageCommandFromAttrs)
	r.Register("mergeCells", newMergeCellsCommandFromAttrs)
	r.Register("updateCell", newUpdateCellCommandFromAttrs)
	r.Register("autoRowHeight", newAutoRowHeightCommandFromAttrs)
	r.Register("repeat", newRepeatCommandFromAttrs)
	r.Register("dataValidation", newDataValidationCommandFromAttrs)
	r.Register("table", newTableCommandFromAttrs)
	r.Register("conditionalFormat", newConditionalFormatCommandFromAttrs)
	r.Register("group", newGroupCommandFromAttrs)
	r.Register("chart", newChartCommandFromAttrs)
	r.Register("definedName", newDefinedNameCommandFromAttrs)
	r.Register("sparkline", newSparklineCommandFromAttrs)
	r.Register("pageBreak", newPageBreakCommandFromAttrs)
	r.Register("autoColWidth", newAutoColWidthCommandFromAttrs)
	r.Register("freezePanes", newFreezePanesCommandFromAttrs)
	r.Register("protect", newProtectCommandFromAttrs)
	r.Register("include", newIncludeCommandFromAttrs)
	return r
}

// Register adds a command factory.
func (r *CommandRegistry) Register(name string, factory CommandFactory) {
	r.factories[name] = factory
}

// Create creates a Command from parsed command data.
// Returns (nil, nil) if the command is not registered. The caller (Filler)
// is responsible for handling unknown commands based on strict mode.
func (r *CommandRegistry) Create(name string, attrs map[string]string) (Command, error) {
	factory, ok := r.factories[name]
	if !ok {
		return nil, nil // unknown command — caller decides how to handle
	}
	return factory(attrs)
}

// KnownNames returns all registered command names in sorted order.
func (r *CommandRegistry) KnownNames() []string {
	names := make([]string, 0, len(r.factories))
	for name := range r.factories {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}
