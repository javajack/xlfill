package xlfill

import (
	"bytes"
	"fmt"
	"io"
	"os"
)

// CompiledTemplate represents a pre-parsed template that can be reused
// for filling with different data sets. The template is parsed once and
// the area/command hierarchy is cached for repeated use.
type CompiledTemplate struct {
	templateBytes []byte
	opts          *Options
	filler        *Filler
}

// Compile parses a template file and returns a reusable CompiledTemplate.
// The template file is read into memory so it can be reopened for each Fill call.
func Compile(templatePath string, opts ...Option) (*CompiledTemplate, error) {
	data, err := os.ReadFile(templatePath)
	if err != nil {
		return nil, NewRuntimeError(fmt.Sprintf("read template %q", templatePath), err)
	}

	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)

	// Validate the template can be parsed by doing a dry-run open + build
	tx, err := filler.openTemplate()
	if err != nil {
		return nil, err
	}
	_, err = filler.BuildAreas(tx)
	tx.Close()
	if err != nil {
		return nil, err
	}

	return &CompiledTemplate{
		templateBytes: data,
		opts:          filler.opts,
		filler:        filler,
	}, nil
}

// Fill processes the compiled template with data and writes to outputPath.
func (ct *CompiledTemplate) Fill(data map[string]any, outputPath string) error {
	out, err := os.Create(outputPath)
	if err != nil {
		return NewRuntimeError(fmt.Sprintf("create output %q", outputPath), err)
	}
	defer out.Close()

	if err := ct.FillWriter(data, out); err != nil {
		os.Remove(outputPath)
		return err
	}
	return nil
}

// FillBytes processes the compiled template with data and returns bytes.
func (ct *CompiledTemplate) FillBytes(data map[string]any) ([]byte, error) {
	var buf bytes.Buffer
	if err := ct.FillWriter(data, &buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// FillBatch processes the compiled template for each item in items,
// producing one output file per item. nameFn is called to determine the
// output file path for each item based on its index and data.
func (ct *CompiledTemplate) FillBatch(
	items []map[string]any,
	nameFn func(index int, data map[string]any) string,
) error {
	for i, data := range items {
		outputPath := nameFn(i, data)
		if err := ct.Fill(data, outputPath); err != nil {
			return fmt.Errorf("batch item %d (%s): %w", i, outputPath, err)
		}
	}
	return nil
}

// FillWriter processes the compiled template with data and writes to w.
// Each call creates a fresh transformer from the cached template bytes,
// re-parses areas (fast since template is in memory), and processes.
func (ct *CompiledTemplate) FillWriter(data map[string]any, w io.Writer) error {
	// Create a filler with the template bytes as reader
	reader := bytes.NewReader(ct.templateBytes)
	filler := NewFiller(WithTemplateReader(reader))
	// Copy over relevant options (all value types are safe to share;
	// slices are defensively copied to allow concurrent FillWriter calls)
	filler.opts.notationBegin = ct.opts.notationBegin
	filler.opts.notationEnd = ct.opts.notationEnd
	filler.opts.clearTemplateCells = ct.opts.clearTemplateCells
	filler.opts.keepTemplateSheet = ct.opts.keepTemplateSheet
	filler.opts.hideTemplateSheet = ct.opts.hideTemplateSheet
	filler.opts.recalculateOnOpen = ct.opts.recalculateOnOpen
	filler.opts.preWrite = ct.opts.preWrite
	filler.opts.strictMode = ct.opts.strictMode
	filler.opts.debugWriter = ct.opts.debugWriter
	filler.opts.progressFunc = ct.opts.progressFunc
	filler.opts.ctx = ct.opts.ctx
	filler.opts.streaming = ct.opts.streaming
	filler.opts.parallelism = ct.opts.parallelism
	filler.opts.autoMode = ct.opts.autoMode
	filler.opts.autoModeHint = ct.opts.autoModeHint
	filler.opts.customFunctions = ct.opts.customFunctions
	filler.opts.i18nBundle = ct.opts.i18nBundle
	filler.opts.docProperties = ct.opts.docProperties
	// Defensive copy of slices so concurrent fills don't share state
	if len(ct.opts.streamingSheets) > 0 {
		filler.opts.streamingSheets = make([]string, len(ct.opts.streamingSheets))
		copy(filler.opts.streamingSheets, ct.opts.streamingSheets)
	}
	if len(ct.opts.areaListeners) > 0 {
		filler.opts.areaListeners = make([]AreaListener, len(ct.opts.areaListeners))
		copy(filler.opts.areaListeners, ct.opts.areaListeners)
	}
	// Copy custom commands
	for name, factory := range ct.opts.customCommands {
		filler.registry.Register(name, factory)
	}

	return filler.FillWriter(data, w)
}
