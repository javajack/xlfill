---
title: "jx:image"
description: Insert an image into the spreadsheet from byte data.
---

`jx:image` inserts an image from byte data into the specified cell area. Use it for employee photos, company logos, product images, chart screenshots, or any visual content.

## Syntax

```
jx:image(src="employee.Photo" imageType="PNG" lastCell="C5")
```

## Attributes

| Attribute | Description | Default | Required |
|-----------|-------------|---------|----------|
| `src` | Expression for image bytes (`[]byte`) | — | Yes |
| `imageType` | Format: `PNG`, `JPEG`, `GIF`, etc. | — | Yes |
| `lastCell` | Bottom-right cell defining the image area | — | Yes |
| `scaleX` | Horizontal scale factor | 1.0 | No |
| `scaleY` | Vertical scale factor | 1.0 | No |

## Example

```go
logoBytes, _ := os.ReadFile("logo.png")

data := map[string]any{
    "company": map[string]any{
        "Name": "Acme Corp",
        "Logo": logoBytes,
    },
}
```

Template cell A1 comment:
```
jx:image(src="company.Logo" imageType="PNG" lastCell="C3")
```

The image fills the A1:C3 area. The cell dimensions control the image size.

## Scaling

Adjust the image size relative to the cell area:

```
jx:image(src="logo" imageType="PNG" scaleX="0.5" scaleY="0.5" lastCell="B2")
```

## Inside loops

Combine with `jx:each` for a different image per row — employee photos, product thumbnails, etc.:

```
Cell A1 comment:
  jx:area(lastCell="D3")
  jx:each(items="employees" var="e" lastCell="D3")

Cell A1 also has:
  jx:image(src="e.Photo" imageType="JPEG" lastCell="A3")
```

Each employee gets their photo in column A, with name/details in columns B-D.

## Common pitfalls

- **`imageType` must match the actual bytes.** Setting `imageType="PNG"` on a JPEG file produces a corrupted image. Sniff the format if your data is mixed: PNG starts with `89 50 4E 47`, JPEG with `FF D8 FF`.
- **The cell area determines image size — but only loosely.** Excel anchors the image to the top-left cell; resizing the cell after the fact rescales the image. Use `scaleX`/`scaleY` if you need a different rendered size than the area suggests.
- **Empty `[]byte` → blank cell.** XLFill silently skips zero-length image data rather than erroring. Validate your byte slice before passing it.
- **Streaming mode rejects images.** Adding an image on a streamed sheet returns an error. Use sequential mode or `WithStreamingSheets` to exclude the image sheet.

## Try it

Download the runnable example: **template** [t10.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t10.xlsx) | **output** [10_image.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/10_image.xlsx) | [code snippet](/xlfill/reference/examples/#10-embed-image)

## Next command

Need to merge cells for section headers?

**[jx:mergeCells &rarr;](/xlfill/commands/mergecells/)**
