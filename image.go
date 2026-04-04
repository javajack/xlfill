package xlfill

import (
	"fmt"
	"strings"
)

// ImageCommand implements the jx:image command for embedding images.
type ImageCommand struct {
	Src       string  // expression returning []byte
	ImageType string  // PNG, JPEG, etc. (default: PNG)
	ScaleX    float64 // width scale (default: 1.0)
	ScaleY    float64 // height scale (default: 1.0)
}

func (c *ImageCommand) Name() string { return "image" }
func (c *ImageCommand) Reset()       {}

// newImageCommandFromAttrs creates an ImageCommand from parsed attributes.
func newImageCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &ImageCommand{
		Src:       attrs["src"],
		ImageType: strings.ToUpper(attrs["imageType"]),
		ScaleX:    1.0,
		ScaleY:    1.0,
	}
	if cmd.Src == "" {
		return nil, fmt.Errorf("image command requires 'src' attribute")
	}
	if cmd.ImageType == "" {
		cmd.ImageType = "PNG"
	}
	// Parse scale values if present
	if s := attrs["scaleX"]; s != "" {
		if _, err := fmt.Sscanf(s, "%f", &cmd.ScaleX); err != nil {
			return nil, fmt.Errorf("invalid scaleX %q: %w", s, err)
		}
	}
	if s := attrs["scaleY"]; s != "" {
		if _, err := fmt.Sscanf(s, "%f", &cmd.ScaleY); err != nil {
			return nil, fmt.Errorf("invalid scaleY %q: %w", s, err)
		}
	}
	return cmd, nil
}

// ApplyAt inserts the image at the given target cell.
func (c *ImageCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	// Evaluate src expression to get image bytes
	val, err := ctx.Evaluate(c.Src)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate image src %q: %w", c.Src, err)
	}

	if val == nil {
		return Size{Width: 1, Height: 1}, nil // skip gracefully
	}

	imgBytes, ok := val.([]byte)
	if !ok {
		return ZeroSize, fmt.Errorf("image src must be []byte, got %T", val)
	}

	cellName := cellRef.CellName()
	if err := transformer.AddImage(cellRef.Sheet, cellName, imgBytes, c.ImageType, c.ScaleX, c.ScaleY); err != nil {
		return ZeroSize, fmt.Errorf("add image at %s: %w", cellRef, err)
	}

	return Size{Width: 1, Height: 1}, nil
}
