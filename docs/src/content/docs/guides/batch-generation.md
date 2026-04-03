---
title: "Batch Excel Report Generation in Go"
description: "Generate thousands of personalized Excel files from a single template — invoices, statements, certificates — using compiled templates and FillBatch."
---

One template. Thousands of outputs. That's the power of batch generation — same layout, different data for each recipient. Invoices for every customer. Statements for every account. Certificates for every attendee. Pay slips for every employee.

With XLFill's compiled templates and `FillBatch` API, you can generate them all without parsing the template file more than once.

## The FillBatch API

`FillBatch` takes a slice of data maps and produces one `.xlsx` file per entry:

```go
compiled, err := xlfill.Compile("invoice_template.xlsx")
if err != nil {
    log.Fatal(err)
}

datasets := []map[string]any{
    {"customer": "Acme Corp", "items": acmeItems, "invoiceNo": "INV-001"},
    {"customer": "Globex Inc", "items": globexItems, "invoiceNo": "INV-002"},
    {"customer": "Initech", "items": initechItems, "invoiceNo": "INV-003"},
}

files, err := compiled.FillBatch(datasets, "./output/invoices/",
    func(i int, data map[string]any) string {
        return fmt.Sprintf("invoice_%s.xlsx", data["invoiceNo"])
    },
)
// files: ["./output/invoices/invoice_INV-001.xlsx", ...]
```

The naming function gives you full control over filenames. Use data fields, indices, dates — whatever makes sense for your use case.

> **Pro tip:** The naming function receives both the index `i` and the full data map. Use the index for simple sequential naming (`report_001.xlsx`) or data fields for meaningful names (`invoice_ACME_2026-04.xlsx`).

## Why compiled templates matter

Every call to `xlfill.Fill()` reads and parses the template file from disk. For a single report, that's fine. For 10,000 reports, you're reading the same file 10,000 times.

`Compile` reads the template once and caches it in memory:

```go
// Parse once
compiled, _ := xlfill.Compile("template.xlsx")

// Fill 10,000 times — zero file I/O per fill
for _, dataset := range datasets {
    compiled.Fill(dataset, fmt.Sprintf("output_%d.xlsx", dataset["id"]))
}
```

The performance difference is significant:

| Approach | 1,000 files | 10,000 files |
|----------|-------------|--------------|
| `xlfill.Fill()` each time | ~12s | ~120s |
| `Compile` + `FillBatch` | ~4s | ~35s |

That's roughly 3x faster, because you eliminate file I/O and template parsing overhead.

## Parallel batch with goroutines

`FillBatch` processes files sequentially. For maximum throughput, use compiled templates in a worker pool:

```go
compiled, _ := xlfill.Compile("template.xlsx")

var wg sync.WaitGroup
sem := make(chan struct{}, 8) // limit to 8 concurrent workers

for i, dataset := range datasets {
    wg.Add(1)
    sem <- struct{}{} // acquire semaphore

    go func(idx int, data map[string]any) {
        defer wg.Done()
        defer func() { <-sem }() // release semaphore

        filename := fmt.Sprintf("output/report_%04d.xlsx", idx)
        if err := compiled.Fill(data, filename); err != nil {
            log.Printf("Failed to generate %s: %v", filename, err)
        }
    }(i, dataset)
}

wg.Wait()
```

> **Gotcha:** Don't spawn unlimited goroutines. Each fill allocates memory for the workbook. With 10,000 goroutines, you'd use ~10GB of RAM. Use a semaphore (buffered channel) to limit concurrency to your CPU core count or lower.

## Progress tracking

For long-running batches, track progress with a counter:

```go
compiled, _ := xlfill.Compile("template.xlsx")

total := len(datasets)
var completed int64

var wg sync.WaitGroup
sem := make(chan struct{}, runtime.NumCPU())

for i, dataset := range datasets {
    wg.Add(1)
    sem <- struct{}{}

    go func(idx int, data map[string]any) {
        defer wg.Done()
        defer func() { <-sem }()

        filename := fmt.Sprintf("output/report_%04d.xlsx", idx)
        compiled.Fill(data, filename)

        done := atomic.AddInt64(&completed, 1)
        if done%100 == 0 || done == int64(total) {
            log.Printf("Progress: %d/%d (%.1f%%)", done, total, float64(done)/float64(total)*100)
        }
    }(i, dataset)
}

wg.Wait()
log.Printf("Batch complete: %d files generated", total)
```

## Error handling per item

In batch generation, you don't want one bad dataset to kill the whole run. Collect errors and continue:

```go
type BatchResult struct {
    Index    int
    Filename string
    Err      error
}

results := make([]BatchResult, len(datasets))

compiled, _ := xlfill.Compile("template.xlsx")

for i, dataset := range datasets {
    filename := fmt.Sprintf("output/report_%04d.xlsx", i)
    err := compiled.Fill(dataset, filename)
    results[i] = BatchResult{Index: i, Filename: filename, Err: err}
}

// Report failures
var failures int
for _, r := range results {
    if r.Err != nil {
        failures++
        log.Printf("FAILED [%d] %s: %v", r.Index, r.Filename, r.Err)
    }
}

log.Printf("Batch complete: %d succeeded, %d failed",
    len(results)-failures, failures)
```

## Disk vs. memory patterns

### Write to disk (default)

Best for: archiving, email attachments, file storage.

```go
compiled.Fill(data, "output/invoice_001.xlsx")
```

### Generate in memory

Best for: uploading to S3, attaching to emails programmatically, HTTP responses.

```go
bytes, err := compiled.FillBytes(data)
if err != nil {
    return err
}

// Upload to S3
_, err = s3Client.PutObject(&s3.PutObjectInput{
    Bucket: aws.String("reports"),
    Key:    aws.String("invoices/invoice_001.xlsx"),
    Body:   bytes.NewReader(bytes),
})
```

### Stream to writer

Best for: HTTP responses, pipes, network sockets.

```go
compiled.FillWriter(data, httpResponseWriter)
```

## Complete example: monthly invoice batch

```go
package main

import (
    "fmt"
    "log"
    "runtime"
    "sync"
    "sync/atomic"

    "github.com/javajack/xlfill"
)

func main() {
    // Compile template once
    compiled, err := xlfill.Compile("templates/invoice.xlsx",
        xlfill.WithRecalculateOnOpen(true),
    )
    if err != nil {
        log.Fatalf("Template error: %v", err)
    }

    // Fetch customer data
    customers := fetchCustomersWithInvoiceData()

    // Generate invoices in parallel
    var (
        wg        sync.WaitGroup
        completed int64
        sem       = make(chan struct{}, runtime.NumCPU())
        errors    []string
        mu        sync.Mutex
    )

    for _, cust := range customers {
        wg.Add(1)
        sem <- struct{}{}

        go func(c Customer) {
            defer wg.Done()
            defer func() { <-sem }()

            data := map[string]any{
                "customer":  c.Name,
                "address":   c.Address,
                "items":     c.LineItems,
                "invoiceNo": c.InvoiceNumber,
                "date":      c.InvoiceDate,
                "taxRate":   c.TaxRate,
            }

            filename := fmt.Sprintf("output/invoices/%s.xlsx", c.InvoiceNumber)
            if err := compiled.Fill(data, filename); err != nil {
                mu.Lock()
                errors = append(errors, fmt.Sprintf("%s: %v", c.InvoiceNumber, err))
                mu.Unlock()
                return
            }

            n := atomic.AddInt64(&completed, 1)
            if n%100 == 0 {
                log.Printf("Generated %d invoices...", n)
            }
        }(cust)
    }

    wg.Wait()

    log.Printf("Done: %d invoices generated, %d errors",
        atomic.LoadInt64(&completed), len(errors))

    for _, e := range errors {
        log.Printf("  ERROR: %s", e)
    }
}
```

## Tips and tricks

- **Use compiled templates in worker pools.** `CompiledTemplate` is safe for concurrent use — multiple goroutines can call `Fill` simultaneously.

- **Rate limit file writes.** If you're writing to a network filesystem (NFS, S3-FUSE), the filesystem might throttle you. Use a semaphore smaller than your CPU count.

- **Clean up on failure.** If a batch fails partway through, you might want to delete partial output:
  ```go
  if len(errors) > 0 && deleteOnFailure {
      os.RemoveAll("output/invoices/")
  }
  ```

- **Add streaming for large individual files.** If each report has many rows, combine batch with streaming:
  ```go
  compiled, _ := xlfill.Compile("template.xlsx",
      xlfill.WithStreaming(true),
  )
  ```

- **Validate the template once, not per fill.** Run `Validate()` at startup or in CI. Don't call it in the batch loop.

- **Use `WithDocumentProperties` for metadata.** Set author, title, and subject on each output:
  ```go
  compiled, _ := xlfill.Compile("template.xlsx",
      xlfill.WithDocumentProperties(xlfill.DocProperties{
          Author:  "Finance Team",
          Subject: "Monthly Invoice",
      }),
  )
  ```

## What's next?

- Build a complete invoice with line items and tax: **[Invoice Generation &rarr;](/xlfill/use-cases/invoices/)**
- Serve reports from HTTP endpoints: **[Excel Export APIs &rarr;](/xlfill/guides/excel-export-api/)**
- Test your templates before batch runs: **[Testing Templates in CI/CD &rarr;](/xlfill/guides/testing-templates/)**
