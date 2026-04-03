---
title: "Building Excel Export APIs in Go"
description: "Serve Excel reports from HTTP endpoints with zero temp files — using XLFill's built-in HTTP handler, streaming, and context cancellation."
---

Every web application eventually needs a "Download as Excel" button. With XLFill, you can serve formatted `.xlsx` files directly from your HTTP handlers — no temp files, no background jobs, no file cleanup. The report streams straight to the browser.

## The one-liner: HTTPHandler

XLFill has a built-in `http.Handler` that does everything for you:

```go
http.Handle("/report", xlfill.HTTPHandler("template.xlsx",
    func(r *http.Request) (map[string]any, error) {
        dept := r.URL.Query().Get("dept")
        employees, err := db.GetEmployeesByDept(dept)
        if err != nil {
            return nil, err
        }
        return map[string]any{"employees": employees}, nil
    },
))
```

That's it. `HTTPHandler` sets the correct `Content-Type` header (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`), adds a `Content-Disposition: attachment` header, and streams the filled template directly to the response. No temp file touches the disk.

> **Pro tip:** `HTTPHandler` returns a standard `http.Handler`, so it works with any Go HTTP framework — net/http, Chi, Gin, Echo, Fiber, you name it.

## Custom headers with FillReader

If you need more control over response headers (custom filename, cache headers, etc.), use `FillReader` directly:

```go
func reportHandler(w http.ResponseWriter, r *http.Request) {
    data, err := buildReportData(r)
    if err != nil {
        http.Error(w, "Failed to build report", http.StatusInternalServerError)
        return
    }

    filename := fmt.Sprintf("report_%s.xlsx", time.Now().Format("2006-01-02"))

    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition",
        fmt.Sprintf(`attachment; filename="%s"`, filename))
    w.Header().Set("Cache-Control", "no-store")

    tmpl, err := os.Open("template.xlsx")
    if err != nil {
        http.Error(w, "Template not found", http.StatusInternalServerError)
        return
    }
    defer tmpl.Close()

    if err := xlfill.FillReader(tmpl, w, data); err != nil {
        // Headers already sent — log the error, can't change the response
        log.Printf("Fill error: %v", err)
    }
}
```

> **Gotcha:** Once you start writing to `w`, you can't change response headers or status code. Validate your data and template *before* calling `FillReader`. If there's a chance of failure, use `FillBytes` first, then write the bytes to the response.

## Compiled templates for high-traffic endpoints

If your endpoint gets frequent hits, parsing the template file on every request wastes I/O. Compile the template once at startup:

```go
var reportTemplate *xlfill.CompiledTemplate

func init() {
    var err error
    reportTemplate, err = xlfill.Compile("template.xlsx",
        xlfill.WithStreaming(true),
    )
    if err != nil {
        log.Fatalf("Failed to compile template: %v", err)
    }
}

func reportHandler(w http.ResponseWriter, r *http.Request) {
    data, err := buildReportData(r)
    if err != nil {
        http.Error(w, err.Error(), http.StatusBadRequest)
        return
    }

    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition",
        `attachment; filename="report.xlsx"`)

    reportTemplate.FillWriter(data, w)
}
```

The template bytes are cached in memory. Each request just fills and writes — zero file I/O.

## Integration with popular routers

### Chi

```go
r := chi.NewRouter()

r.Get("/api/reports/employees", func(w http.ResponseWriter, r *http.Request) {
    dept := chi.URLParam(r, "dept")
    data := fetchEmployeeData(dept)

    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition",
        `attachment; filename="employees.xlsx"`)

    tmpl, _ := os.Open("templates/employees.xlsx")
    defer tmpl.Close()
    xlfill.FillReader(tmpl, w, data)
})
```

### Gin

```go
r := gin.Default()

r.GET("/api/reports/employees", func(c *gin.Context) {
    dept := c.Query("dept")
    data := fetchEmployeeData(dept)

    c.Header("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c.Header("Content-Disposition",
        `attachment; filename="employees.xlsx"`)

    tmpl, _ := os.Open("templates/employees.xlsx")
    defer tmpl.Close()
    xlfill.FillReader(tmpl, c.Writer, data)
})
```

### Echo

```go
e := echo.New()

e.GET("/api/reports/employees", func(c echo.Context) error {
    dept := c.QueryParam("dept")
    data := fetchEmployeeData(dept)

    c.Response().Header().Set("Content-Disposition",
        `attachment; filename="employees.xlsx"`)

    tmpl, _ := os.Open("templates/employees.xlsx")
    defer tmpl.Close()
    return xlfill.FillReader(tmpl, c.Response(), data)
})
```

## Context cancellation for timeouts

Large reports can take time. Use `WithContext` to respect request deadlines:

```go
func reportHandler(w http.ResponseWriter, r *http.Request) {
    ctx, cancel := context.WithTimeout(r.Context(), 30*time.Second)
    defer cancel()

    data := buildReportData(r)

    bytes, err := xlfill.FillBytes("template.xlsx", data,
        xlfill.WithContext(ctx),
        xlfill.WithStreaming(true),
    )
    if err != nil {
        if ctx.Err() == context.DeadlineExceeded {
            http.Error(w, "Report generation timed out", http.StatusGatewayTimeout)
            return
        }
        http.Error(w, "Report generation failed", http.StatusInternalServerError)
        return
    }

    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition",
        `attachment; filename="report.xlsx"`)
    w.Header().Set("Content-Length", strconv.Itoa(len(bytes)))
    w.Write(bytes)
}
```

> **Pro tip:** Using `FillBytes` + `Content-Length` is better for download progress bars in browsers. The browser knows the total size upfront.

## Streaming for large responses

For reports with 10K+ rows, enable streaming mode to reduce memory usage and speed up generation:

```go
http.Handle("/large-report", xlfill.HTTPHandler("template.xlsx",
    func(r *http.Request) (map[string]any, error) {
        return fetchLargeDataset()
    },
    xlfill.WithStreaming(true),
))
```

Streaming mode writes rows to the output as they're processed, instead of building the entire workbook in memory first. This means 3x faster generation and 60% less memory for large datasets.

## Progress reporting for long downloads

For reports that take more than a couple of seconds, you might want to track progress (for logging, monitoring, or a progress API):

```go
func reportHandler(w http.ResponseWriter, r *http.Request) {
    data := buildReportData(r)

    var lastProgress xlfill.FillProgress

    bytes, err := xlfill.FillBytes("template.xlsx", data,
        xlfill.WithStreaming(true),
        xlfill.WithProgressFunc(func(p xlfill.FillProgress) {
            lastProgress = p
            log.Printf("Report progress: %d rows in %v",
                p.ProcessedRows, p.Elapsed)
        }),
    )
    if err != nil {
        http.Error(w, err.Error(), http.StatusInternalServerError)
        return
    }

    log.Printf("Report complete: %d rows in %v",
        lastProgress.ProcessedRows, lastProgress.Elapsed)

    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Write(bytes)
}
```

## Error handling patterns

### Validate before filling

Catch data problems before starting the (potentially slow) fill:

```go
func reportHandler(w http.ResponseWriter, r *http.Request) {
    data, err := buildReportData(r)
    if err != nil {
        http.Error(w, "Invalid parameters", http.StatusBadRequest)
        return
    }

    // Quick validation — catches missing variables, type mismatches
    issues, err := xlfill.ValidateData("template.xlsx", data)
    if err != nil {
        http.Error(w, "Template error", http.StatusInternalServerError)
        return
    }
    if len(issues) > 0 {
        http.Error(w, fmt.Sprintf("Data validation: %s", issues[0]), http.StatusBadRequest)
        return
    }

    // Safe to fill
    bytes, _ := xlfill.FillBytes("template.xlsx", data)
    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Write(bytes)
}
```

### Typed error handling

```go
if err := xlfill.Fill("template.xlsx", "output.xlsx", data); err != nil {
    var xlErr *xlfill.XLFillError
    if errors.As(err, &xlErr) {
        switch xlErr.Kind {
        case xlfill.ErrTemplate:
            // Template structure problem — fix the .xlsx file
            log.Printf("Template error at %s: %s", xlErr.Cell, xlErr.Message)
        case xlfill.ErrData:
            // Data doesn't match template expectations
            log.Printf("Data error at %s: %s", xlErr.Cell, xlErr.Message)
        case xlfill.ErrRuntime:
            // I/O or system error
            log.Printf("Runtime error: %s", xlErr.Message)
        }
    }
}
```

## Tips and tricks

- **Content-Disposition filename:** Use RFC 5987 encoding for non-ASCII filenames:
  ```go
  filename := url.PathEscape("rapport_février.xlsx")
  w.Header().Set("Content-Disposition",
      fmt.Sprintf(`attachment; filename*=UTF-8''%s`, filename))
  ```

- **ETags for caching:** If your report data has a known version, generate an ETag to avoid re-downloading unchanged reports:
  ```go
  etag := fmt.Sprintf(`"%x"`, md5.Sum([]byte(dataVersion)))
  w.Header().Set("ETag", etag)
  if r.Header.Get("If-None-Match") == etag {
      w.WriteHeader(http.StatusNotModified)
      return
  }
  ```

- **Gzip middleware compatibility:** `.xlsx` files are already ZIP-compressed. Gzip middleware adds CPU overhead with near-zero size reduction. Exclude `.xlsx` responses from gzip:
  ```go
  // Chi middleware example
  compressor := middleware.NewCompressor(5)
  compressor.SetEncoder("xlsx", func(w io.Writer, level int) io.WriteCloser {
      return nopCloser{w} // pass through, don't compress
  })
  ```

- **Rate limiting:** For expensive reports, add rate limiting per user to prevent abuse. A single 100K-row report might take a few seconds of CPU.

- **Multiple templates:** For apps with many report types, compile all templates at startup and store them in a map:
  ```go
  var templates = map[string]*xlfill.CompiledTemplate{}

  func init() {
      for _, name := range []string{"sales", "inventory", "payroll"} {
          t, _ := xlfill.Compile(fmt.Sprintf("templates/%s.xlsx", name))
          templates[name] = t
      }
  }
  ```

## What's next?

- Generate thousands of reports in one go: **[Batch Generation &rarr;](/xlfill/guides/batch-generation/)**
- Add charts and dashboards to your reports: **[Charts and Dashboards &rarr;](/xlfill/guides/charts-and-dashboards/)**
- Test your templates in CI: **[Testing Templates in CI/CD &rarr;](/xlfill/guides/testing-templates/)**
