package xlfill

import (
	"bytes"
	"fmt"
	"io"
	"os"

	"github.com/xuri/excelize/v2"
)

// Fill processes a template file and writes the populated output to outputPath.
func Fill(templatePath, outputPath string, data map[string]any, opts ...Option) error {
	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.Fill(data, outputPath)
}

// FillBytes processes a template file and returns the populated output as bytes.
func FillBytes(templatePath string, data map[string]any, opts ...Option) ([]byte, error) {
	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.FillBytes(data)
}

// FillReader processes a template from an io.Reader and writes to an io.Writer.
func FillReader(template io.Reader, output io.Writer, data map[string]any, opts ...Option) error {
	allOpts := append([]Option{WithTemplateReader(template)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.FillWriter(data, output)
}

// Fill processes the template with data and writes to outputPath.
func (f *Filler) Fill(data map[string]any, outputPath string) error {
	out, err := os.Create(outputPath)
	if err != nil {
		return NewRuntimeError(fmt.Sprintf("create output file %q", outputPath), err)
	}
	defer out.Close()

	if err := f.FillWriter(data, out); err != nil {
		os.Remove(outputPath)
		return err
	}
	return nil
}

// FillBytes processes the template with data and returns the output as bytes.
func (f *Filler) FillBytes(data map[string]any) ([]byte, error) {
	var buf bytes.Buffer
	if err := f.FillWriter(data, &buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// FillWriter processes the template with data and writes to w.
func (f *Filler) FillWriter(data map[string]any, w io.Writer) error {
	// Open template
	etx, err := f.openTemplate()
	if err != nil {
		return err
	}
	defer etx.Close()

	// Auto-mode: analyze template and select optimal strategy.
	// Uses LOCAL variables to avoid mutating f.opts (safe for concurrent Filler reuse).
	autoStreaming := f.opts.streaming
	autoParallelism := f.opts.parallelism
	if f.opts.autoMode {
		suggestion, sugErr := f.suggestModeFromTransformer(etx, f.opts.autoModeHint)
		if sugErr == nil {
			switch suggestion.Mode {
			case ModeStreaming:
				autoStreaming = true
				autoParallelism = 0
			case ModeParallel:
				autoStreaming = false
				autoParallelism = suggestion.Parallelism
			default:
				autoStreaming = false
				autoParallelism = 0
			}
		}
		// On analysis error, fall through to sequential
	}

	// If auto-mode ran BuildAreas for analysis, reset target tracking
	// to avoid duplicate entries when BuildAreas runs again below.
	if f.opts.autoMode {
		etx.ResetTargetCellRefs()
	}

	// Determine the active transformer based on options.
	// Streaming and parallel are mutually exclusive; parallel takes precedence.
	useStreaming := autoStreaming && autoParallelism <= 1
	useParallel := autoParallelism > 1

	var tx Transformer = etx
	var stx *StreamingTransformer

	if useStreaming {
		sheets := etx.GetSheetNames()
		if len(sheets) > 0 {
			targetSheet := sheets[0]
			shouldStream := len(f.opts.streamingSheets) == 0 // stream all if no specific sheets
			for _, s := range f.opts.streamingSheets {
				if s == targetSheet {
					shouldStream = true
					break
				}
			}
			if shouldStream {
				var serr error
				stx, serr = NewStreamingTransformer(etx, targetSheet)
				if serr != nil {
					return fmt.Errorf("init streaming: %w", serr)
				}
				tx = stx
			}
		}
	}

	if useParallel {
		tx = NewConcurrentTransformer(tx)
	}

	// Create context
	ctxOpts := []ContextOption{}
	if f.opts.notationBegin != "${" || f.opts.notationEnd != "}" {
		ctxOpts = append(ctxOpts, WithNotation(f.opts.notationBegin, f.opts.notationEnd))
	}
	if len(f.opts.customFunctions) > 0 {
		ctxOpts = append(ctxOpts, WithCustomFunctions(f.opts.customFunctions))
	}
	if len(f.opts.i18nBundle) > 0 {
		ctxOpts = append(ctxOpts, WithI18nBundle(f.opts.i18nBundle))
	}
	ctx := NewContext(data, ctxOpts...)

	// Build areas from template comments
	areas, err := f.BuildAreas(tx)
	if err != nil {
		return err
	}

	// Propagate parallelism setting to areas
	if useParallel {
		propagateParallelism(areas, autoParallelism)
	}

	// Debug trace start
	if f.debug != nil {
		f.debug.TraceArea(areas[0], areas[0].StartCell)
	}

	// Process each area
	for _, area := range areas {
		if _, err := area.ApplyAt(area.StartCell, ctx); err != nil {
			return fmt.Errorf("process area at %s: %w", area.StartCell, err)
		}

		// Clear template cells if configured
		if f.opts.clearTemplateCells {
			area.clearTemplateCells(ctx)
		}
	}

	// Debug trace done
	if f.debug != nil {
		f.debug.TraceDone()
	}

	// Execute deferred actions (e.g., jx:table, jx:chart that need final output ranges)
	for _, action := range ctx.deferred.Actions() {
		if action.Execute != nil {
			if f.debug != nil {
				f.debug.TraceDeferredAction(action.Name, action.Sheet, action.StartRow, action.EndRow)
			}
			if err := action.Execute(etx); err != nil {
				return fmt.Errorf("deferred action %q: %w", action.Name, err)
			}
		}
	}

	// Recalculate formulas on open
	if f.opts.recalculateOnOpen {
		if err := etx.SetRecalculateOnOpen(true); err != nil {
			return fmt.Errorf("set recalculate on open: %w", err)
		}
	}

	// Set document properties if configured
	if f.opts.docProperties != nil {
		if err := etx.file.SetDocProps(f.opts.docProperties); err != nil {
			return fmt.Errorf("set document properties: %w", err)
		}
	}

	// Pre-write callback (always gets the underlying ExcelizeTransformer)
	if f.opts.preWrite != nil {
		if err := f.opts.preWrite(etx); err != nil {
			return fmt.Errorf("pre-write callback: %w", err)
		}
	}

	// Write output
	return tx.Write(w)
}

// propagateParallelism sets the parallelism field on all areas recursively.
func propagateParallelism(areas []*Area, n int) {
	for _, area := range areas {
		area.parallelism = n
		for _, b := range area.Bindings {
			if childArea := getCommandArea(b.Command); childArea != nil {
				propagateParallelism([]*Area{childArea}, n)
			}
		}
	}
}

// ValidateData checks if the provided data satisfies the template's requirements.
// It parses all expressions in the template and verifies that the data map contains
// the required top-level keys. Returns validation issues for mismatches.
func ValidateData(templatePath string, data map[string]any, opts ...Option) ([]ValidationIssue, error) {
	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.ValidateData(data)
}

// openTemplate opens the template from file path or reader.
func (f *Filler) openTemplate() (*ExcelizeTransformer, error) {
	if f.opts.templateReader != nil {
		file, err := excelize.OpenReader(f.opts.templateReader)
		if err != nil {
			return nil, NewRuntimeError("open template reader", err)
		}
		return NewExcelizeTransformer(file)
	}
	if f.opts.templatePath != "" {
		return OpenTemplate(f.opts.templatePath)
	}
	return nil, NewRuntimeError("no template specified: use WithTemplate or WithTemplateReader", nil)
}

// clearTemplateCells clears cells that still contain unexpanded template expressions.
func (a *Area) clearTemplateCells(ctx *Context) {
	// We only clear the source area cells that weren't overwritten by command output.
	// The area's ClearCells method handles this.
	// For now, no-op — the Transform already wrote evaluated values to target cells.
	// Template cells outside any processed area retain their expressions, which is
	// handled by clearing the area source if the output target differs from source.
}
