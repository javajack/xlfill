package goxls

import (
	"image"
	"image/color"
	"image/png"
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// createTestPNG generates a small PNG image for testing.
func createTestPNG(t *testing.T) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, 10, 10))
	for x := 0; x < 10; x++ {
		for y := 0; y < 10; y++ {
			img.Set(x, y, color.RGBA{R: 255, A: 255})
		}
	}
	var buf bytes.Buffer
	require.NoError(t, png.Encode(&buf, img))
	return buf.Bytes()
}

func TestImageCommand_PNG(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	imgBytes := createTestPNG(t)
	ctx := NewContext(map[string]any{"img": imgBytes})

	cmd := &ImageCommand{Src: "img", ImageType: "PNG", ScaleX: 1.0, ScaleY: 1.0}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)

	// Verify the file can be written (image embedded)
	var buf bytes.Buffer
	require.NoError(t, tx.Write(&buf))
	assert.True(t, buf.Len() > 0)
}

func TestImageCommand_NilBytes(t *testing.T) {
	f := excelize.NewFile()
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	ctx := NewContext(map[string]any{"img": nil})
	cmd := &ImageCommand{Src: "img", ImageType: "PNG", ScaleX: 1.0, ScaleY: 1.0}
	size, err := cmd.ApplyAt(NewCellRef("Sheet1", 0, 0), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size) // graceful skip
}

func TestImageCommand_WithScaling(t *testing.T) {
	f := excelize.NewFile()
	sheet := "Sheet1"
	tx, err := NewExcelizeTransformer(f)
	require.NoError(t, err)
	defer tx.Close()

	imgBytes := createTestPNG(t)
	ctx := NewContext(map[string]any{"img": imgBytes})

	cmd := &ImageCommand{Src: "img", ImageType: "PNG", ScaleX: 2.0, ScaleY: 0.5}
	size, err := cmd.ApplyAt(NewCellRef(sheet, 2, 3), ctx, tx)
	require.NoError(t, err)
	assert.Equal(t, Size{Width: 1, Height: 1}, size)
}

func TestNewImageCommandFromAttrs(t *testing.T) {
	cmd, err := newImageCommandFromAttrs(map[string]string{
		"src": "myImg", "imageType": "jpeg", "scaleX": "1.5", "scaleY": "2.0",
	})
	require.NoError(t, err)
	img := cmd.(*ImageCommand)
	assert.Equal(t, "myImg", img.Src)
	assert.Equal(t, "JPEG", img.ImageType)
	assert.Equal(t, 1.5, img.ScaleX)
	assert.Equal(t, 2.0, img.ScaleY)
}

func TestNewImageCommandFromAttrs_MissingSrc(t *testing.T) {
	_, err := newImageCommandFromAttrs(map[string]string{})
	assert.Error(t, err)
}

func TestNewImageCommandFromAttrs_Defaults(t *testing.T) {
	cmd, err := newImageCommandFromAttrs(map[string]string{"src": "img"})
	require.NoError(t, err)
	img := cmd.(*ImageCommand)
	assert.Equal(t, "PNG", img.ImageType)
	assert.Equal(t, 1.0, img.ScaleX)
	assert.Equal(t, 1.0, img.ScaleY)
}
