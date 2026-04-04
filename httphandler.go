package xlfill

import (
	"fmt"
	"net/http"
	"strings"
)

// sanitizeFilename removes characters that could be used for header injection.
func sanitizeFilename(name string) string {
	return strings.Map(func(r rune) rune {
		if r < 32 || r == '"' || r == '\\' || r == '/' || r == ':' || r == ';' {
			return '_'
		}
		return r
	}, name)
}

// HTTPHandler creates an http.HandlerFunc that fills a compiled template
// with data extracted from the request and streams the result as an Excel download.
//
// The dataFn receives the HTTP request and returns:
//   - data: the template data
//   - filename: the suggested download filename (without extension)
//   - err: any error (results in 500 response)
func HTTPHandler(
	compiled *CompiledTemplate,
	dataFn func(r *http.Request) (data map[string]any, filename string, err error),
) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		data, filename, err := dataFn(r)
		if err != nil {
			http.Error(w, fmt.Sprintf("data extraction failed: %v", err), http.StatusInternalServerError)
			return
		}
		if filename == "" {
			filename = "report"
		}
		filename = sanitizeFilename(filename)

		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename="%s.xlsx"`, filename))
		w.Header().Set("X-Content-Type-Options", "nosniff")
		w.Header().Set("Cache-Control", "no-store")

		if err := compiled.FillWriter(data, w); err != nil {
			// Headers already sent, can't change status code reliably.
			// Caller should handle via middleware or logging.
			return
		}
	}
}
