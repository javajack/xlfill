package xlfill

import (
	"fmt"
	"net/http"
)

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

		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename="%s.xlsx"`, filename))

		if err := compiled.FillWriter(data, w); err != nil {
			// Headers already sent, can't change status code reliably.
			// Caller should handle via middleware or logging.
			return
		}
	}
}
