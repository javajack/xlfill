package xlfill

import (
	"context"
	"io"
	"time"
)

// FillProgress reports progress during template filling.
type FillProgress struct {
	ProcessedRows int
	TotalRows     int // estimated, may be 0 if unknown
	CurrentSheet  string
	Elapsed       time.Duration
}

// ProgressFunc is called periodically during Fill to report progress.
type ProgressFunc func(FillProgress)

// Options holds configuration for the Filler.
type Options struct {
	templatePath       string
	templateReader     io.Reader
	notationBegin      string
	notationEnd        string
	customCommands     map[string]CommandFactory
	clearTemplateCells bool
	keepTemplateSheet  bool
	hideTemplateSheet  bool
	recalculateOnOpen  bool
	areaListeners      []AreaListener
	preWrite           func(Transformer) error

	// Enhancement options
	strictMode   bool         // error on unknown commands
	debugWriter  io.Writer    // trace output writer
	progressFunc ProgressFunc // progress callback
	ctx          context.Context
	streaming    bool // use StreamWriter for output
	parallelism  int  // number of goroutines for parallel each (0 = sequential)
	autoMode     bool // auto-detect optimal mode from template structure
	autoModeHint map[string]any
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

// WithStrictMode enables strict mode where unknown jx: commands produce errors
// instead of being silently ignored. When disabled (default), unknown commands
// are collected as warnings accessible via Filler.Warnings().
func WithStrictMode(strict bool) Option {
	return func(o *Options) { o.strictMode = strict }
}

// WithDebugWriter enables trace output during template processing.
// The trace includes area processing, command execution, expression evaluation,
// and timing information — useful for debugging templates and performance tuning.
func WithDebugWriter(w io.Writer) Option {
	return func(o *Options) { o.debugWriter = w }
}

// WithProgressFunc sets a callback for progress reporting during Fill.
// Called periodically with the number of processed rows and elapsed time.
func WithProgressFunc(fn ProgressFunc) Option {
	return func(o *Options) { o.progressFunc = fn }
}

// WithContext sets a context for cancellation support.
// If the context is cancelled during processing, Fill returns the context error.
func WithContext(ctx context.Context) Option {
	return func(o *Options) { o.ctx = ctx }
}

// WithStreaming enables streaming mode for large outputs.
// In streaming mode, output rows are flushed incrementally via excelize StreamWriter
// instead of holding the entire workbook in memory. This reduces memory usage for
// outputs with 100K+ rows.
//
// Limitations: formula post-processing (reference remapping), hyperlinks, and images
// are not supported in streaming mode. Streaming and parallel are mutually exclusive;
// if both are set, parallel takes precedence.
func WithStreaming(enabled bool) Option {
	return func(o *Options) { o.streaming = enabled }
}

// WithParallelism sets the number of goroutines for parallel jx:each processing.
// Only jx:each commands with direction="DOWN" and fixed-height areas (no nested
// each/repeat) are parallelized. Other commands fall back to sequential.
// Set to 0 (default) or 1 to disable parallelism.
//
// When enabled, the Transformer is wrapped in a ConcurrentTransformer for thread-safe
// writes, and each goroutine gets an independent Context clone.
func WithParallelism(n int) Option {
	return func(o *Options) { o.parallelism = n }
}
