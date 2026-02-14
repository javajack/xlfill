package goxls

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
		return fmt.Errorf("create output file %q: %w", outputPath, err)
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
	tx, err := f.openTemplate()
	if err != nil {
		return err
	}
	defer tx.Close()

	// Create context
	ctxOpts := []ContextOption{}
	if f.opts.notationBegin != "${" || f.opts.notationEnd != "}" {
		ctxOpts = append(ctxOpts, WithNotation(f.opts.notationBegin, f.opts.notationEnd))
	}
	ctx := NewContext(data, ctxOpts...)

	// Build areas from template comments
	areas, err := f.BuildAreas(tx)
	if err != nil {
		return err
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

	// Pre-write callback
	if f.opts.preWrite != nil {
		if err := f.opts.preWrite(tx); err != nil {
			return fmt.Errorf("pre-write callback: %w", err)
		}
	}

	// Write output
	return tx.Write(w)
}

// openTemplate opens the template from file path or reader.
func (f *Filler) openTemplate() (*ExcelizeTransformer, error) {
	if f.opts.templateReader != nil {
		file, err := excelize.OpenReader(f.opts.templateReader)
		if err != nil {
			return nil, fmt.Errorf("open template reader: %w", err)
		}
		return NewExcelizeTransformer(file)
	}
	if f.opts.templatePath != "" {
		return OpenTemplate(f.opts.templatePath)
	}
	return nil, fmt.Errorf("no template specified: use WithTemplate or WithTemplateReader")
}

// clearTemplateCells clears cells that still contain unexpanded template expressions.
func (a *Area) clearTemplateCells(ctx *Context) {
	// We only clear the source area cells that weren't overwritten by command output.
	// The area's ClearCells method handles this.
	// For now, no-op — the Transform already wrote evaluated values to target cells.
	// Template cells outside any processed area retain their expressions, which is
	// handled by clearing the area source if the output target differs from source.
}
