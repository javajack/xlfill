package xlfill

import (
	"context"
	"io"
	"time"

	"github.com/xuri/excelize/v2"
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

	// Custom template functions registered via WithFunction.
	customFunctions map[string]any

	// i18n resource bundle for the t() template function.
	i18nBundle map[string]string

	// Document properties to set on the output workbook.
	docProperties *excelize.DocProperties

	// Selective sheet streaming: when non-empty, only these sheets use StreamWriter.
	streamingSheets []string
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

// MaxParallelism caps the maximum number of parallel goroutines to prevent resource exhaustion.
const MaxParallelism = 256

// WithParallelism sets the number of goroutines for parallel jx:each processing.
// Only jx:each commands with direction="DOWN" and fixed-height areas (no nested
// each/repeat) are parallelized. Other commands fall back to sequential.
// Set to 0 (default) or 1 to disable parallelism.
//
// When enabled, the Transformer is wrapped in a ConcurrentTransformer for thread-safe
// writes, and each goroutine gets an independent Context clone.
func WithParallelism(n int) Option {
	return func(o *Options) {
		if n > MaxParallelism {
			n = MaxParallelism
		}
		if n < 0 {
			n = 0
		}
		o.parallelism = n
	}
}

// WithFunction registers a custom template function that will be available
// in all template expressions. The function name must not collide with existing
// data keys (data keys take precedence).
//
// Example:
//
//	filler := NewFiller(
//	    WithTemplate("template.xlsx"),
//	    WithFunction("currency", func(amount float64) string {
//	        return fmt.Sprintf("$%.2f", amount)
//	    }),
//	)
func WithFunction(name string, fn any) Option {
	return func(o *Options) {
		if o.customFunctions == nil {
			o.customFunctions = make(map[string]any)
		}
		o.customFunctions[name] = fn
	}
}

// WithI18n registers an i18n resource bundle for the t() template function.
// The bundle maps translation keys to localized strings.
//
// Example:
//
//	filler := NewFiller(
//	    WithTemplate("template.xlsx"),
//	    WithI18n(map[string]string{
//	        "greeting": "Hola",
//	        "farewell": "Adios",
//	    }),
//	)
//
// In the template: ${t("greeting")} renders as "Hola".
func WithI18n(bundle map[string]string) Option {
	return func(o *Options) { o.i18nBundle = bundle }
}

// WithDocumentProperties sets Excel document core properties (title, creator, etc.)
// on the output workbook. Supported keys: "title", "subject", "creator",
// "description", "keywords", "category", "language", "version",
// "lastModifiedBy", "contentStatus", "revision", "identifier".
func WithDocumentProperties(props map[string]string) Option {
	return func(o *Options) {
		dp := &excelize.DocProperties{}
		if v, ok := props["title"]; ok {
			dp.Title = v
		}
		if v, ok := props["subject"]; ok {
			dp.Subject = v
		}
		if v, ok := props["creator"]; ok {
			dp.Creator = v
		}
		if v, ok := props["description"]; ok {
			dp.Description = v
		}
		if v, ok := props["keywords"]; ok {
			dp.Keywords = v
		}
		if v, ok := props["category"]; ok {
			dp.Category = v
		}
		if v, ok := props["language"]; ok {
			dp.Language = v
		}
		if v, ok := props["version"]; ok {
			dp.Version = v
		}
		if v, ok := props["lastModifiedBy"]; ok {
			dp.LastModifiedBy = v
		}
		if v, ok := props["contentStatus"]; ok {
			dp.ContentStatus = v
		}
		if v, ok := props["revision"]; ok {
			dp.Revision = v
		}
		if v, ok := props["identifier"]; ok {
			dp.Identifier = v
		}
		o.docProperties = dp
	}
}

// WithStreamingSheets enables streaming mode for specific sheets only.
// When set, only the named sheets use excelize StreamWriter; other sheets
// are processed normally. This implicitly enables streaming mode.
func WithStreamingSheets(sheets ...string) Option {
	return func(o *Options) {
		o.streaming = true
		o.streamingSheets = sheets
	}
}
