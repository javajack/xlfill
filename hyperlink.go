package xlfill

// HyperlinkValue represents a clickable hyperlink in a cell.
// When an expression evaluates to this type, the transformer writes both
// the display text and the hyperlink URL.
type HyperlinkValue struct {
	URL     string
	Display string
}

// String returns the display text for the hyperlink.
func (h HyperlinkValue) String() string {
	if h.Display != "" {
		return h.Display
	}
	return h.URL
}

// Hyperlink creates a HyperlinkValue for use in template expressions.
// Usage in template: ${hyperlink(e.URL, e.Title)}
func Hyperlink(url, display string) HyperlinkValue {
	return HyperlinkValue{URL: url, Display: display}
}
