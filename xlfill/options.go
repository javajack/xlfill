package xlfill

import "io"

// Options holds configuration for the Filler.
type Options struct {
	templatePath        string
	templateReader      io.Reader
	notationBegin       string
	notationEnd         string
	customCommands      map[string]CommandFactory
	clearTemplateCells  bool
	keepTemplateSheet   bool
	hideTemplateSheet   bool
	recalculateOnOpen   bool
	areaListeners       []AreaListener
	preWrite            func(Transformer) error
}

func defaultOptions() *Options {
	return &Options{
		notationBegin:      "${",
		notationEnd:        "}",
		clearTemplateCells: true,
	}
}

// Option configures the Filler.
type Option func(*Options)

// WithTemplate sets the template file path.
func WithTemplate(path string) Option {
	return func(o *Options) { o.templatePath = path }
}

// WithTemplateReader sets the template as an io.Reader.
func WithTemplateReader(r io.Reader) Option {
	return func(o *Options) { o.templateReader = r }
}

// WithExpressionNotation sets the expression delimiters (default: "${", "}").
func WithExpressionNotation(begin, end string) Option {
	return func(o *Options) {
		o.notationBegin = begin
		o.notationEnd = end
	}
}

// WithCommand registers a custom command factory.
func WithCommand(name string, factory CommandFactory) Option {
	return func(o *Options) {
		if o.customCommands == nil {
			o.customCommands = make(map[string]CommandFactory)
		}
		o.customCommands[name] = factory
	}
}

// WithClearTemplateCells controls whether template cells are cleared after processing (default: true).
func WithClearTemplateCells(clear bool) Option {
	return func(o *Options) { o.clearTemplateCells = clear }
}

// WithKeepTemplateSheet keeps the original template sheet in the output.
func WithKeepTemplateSheet(keep bool) Option {
	return func(o *Options) { o.keepTemplateSheet = keep }
}

// WithHideTemplateSheet hides the template sheet instead of deleting it.
func WithHideTemplateSheet(hide bool) Option {
	return func(o *Options) { o.hideTemplateSheet = hide }
}

// WithRecalculateOnOpen tells Excel to recalculate all formulas when the file is opened.
func WithRecalculateOnOpen(recalc bool) Option {
	return func(o *Options) { o.recalculateOnOpen = recalc }
}

// WithAreaListener adds a listener that is notified before/after each cell transformation.
func WithAreaListener(listener AreaListener) Option {
	return func(o *Options) { o.areaListeners = append(o.areaListeners, listener) }
}

// WithPreWrite sets a callback executed before writing the output.
func WithPreWrite(fn func(Transformer) error) Option {
	return func(o *Options) { o.preWrite = fn }
}
