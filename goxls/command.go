package goxls

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
	return r
}

// Register adds a command factory.
func (r *CommandRegistry) Register(name string, factory CommandFactory) {
	r.factories[name] = factory
}

// Create creates a Command from parsed command data.
func (r *CommandRegistry) Create(name string, attrs map[string]string) (Command, error) {
	factory, ok := r.factories[name]
	if !ok {
		return nil, nil // unknown commands are silently ignored
	}
	return factory(attrs)
}
