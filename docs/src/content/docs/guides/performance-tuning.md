---
title: Performance Tuning
description: Streaming mode, parallel processing, compiled templates, and auto-mode — optimize XLFill for any workload.
---

XLFill ships with three processing modes and a compiled template API. Choose the right combination for your workload, or let XLFill decide automatically.

## The three modes

| Mode | Best for | Speedup | Memory savings | Tradeoffs |
|------|----------|---------|----------------|-----------|
| **Sequential** (default) | < 10K rows | Baseline | Baseline | None — full feature support |
| **Streaming** | > 10K rows, simple templates | **3x faster** | **60% less memory** | No formula remapping, no images, no hyperlinks |
| **Parallel** | > 100 rows, CPU-bound expressions | Scales with cores | Similar to sequential | Fixed-height areas only, mutex overhead |

### Streaming mode

Streaming writes output rows incrementally via excelize's StreamWriter instead of holding the entire workbook in memory. Ideal for large, formula-free data exports.

```go
xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithStreaming(true),
)
```

**Benchmark (1,000 rows):**

| | Sequential | Streaming |
|---|---|---|
| Time | 27.5ms | **8.9ms** |
| Memory | 8.3 MB | **3.3 MB** |
| Allocs | 110K | **43K** |

:::caution[Streaming trades features for speed — read this before enabling]
Streaming mode skips or drops several features. If your template uses any of these, you'll get **silently wrong output** unless you handle it:

- **Hyperlinks are dropped** — the display text lands in the cell but the link is gone. You can detect this by inspecting `Filler.Warnings()` after the fill.
- **Images return an error** when added on a streamed sheet.
- **Formula reference remapping is skipped** — `=SUM(B1:B1)` won't expand to `=SUM(B1:B100)` on a streamed sheet.
- **Per-row height changes are silently ignored** after the StreamWriter starts.

Easy way to avoid surprises: use **`WithAutoMode`** instead of `WithStreaming(true)`. Auto-mode inspects your template and only picks streaming when it's compatible.
:::

**Limitations on streamed sheets:**
- Formula reference remapping is skipped (formulas are written verbatim)
- Hyperlinks are dropped and surfaced via `Filler.Warnings()`
- Images return an error
- Per-row height changes after the StreamWriter starts are silently ignored
- Rows must be written in ascending order (guaranteed by template processing)

#### Selective sheet streaming

If your workbook mixes a huge data sheet with a small summary sheet, stream only the big one:

```go
xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithStreamingSheets("BigData"), // stream this sheet only
)
```

The named sheet uses StreamWriter; other sheets use the in-memory path with full feature support (hyperlinks, images, formula remapping). Pass multiple sheet names to stream more than one:

```go
xlfill.WithStreamingSheets("Sales", "Returns")
```

Sheets named in the list that don't exist in the workbook are skipped with a warning (inspect via `Filler.Warnings()`).

#### Catching streaming warnings

```go
filler := xlfill.NewFiller(
    xlfill.WithTemplate("template.xlsx"),
    xlfill.WithStreaming(true),
)
if err := filler.Fill(data, "report.xlsx"); err != nil {
    log.Fatal(err)
}
for _, w := range filler.Warnings() {
    log.Printf("warning: %s", w)
}
```

### Parallel mode

Parallel mode runs `jx:each` iterations concurrently using goroutines. Each goroutine gets an independent context clone and a pre-computed row offset.

```go
xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithParallelism(4), // 4 goroutines
)
```

**When it kicks in:**
- Direction must be `DOWN` (column-wise expansion can't be pre-offset)
- Area must be fixed-height (no nested `jx:each` or `jx:repeat` that change output height)
- Item count must be >= the parallelism value
- Otherwise, falls back to sequential automatically — no error, no config change needed

**Safety guarantees:**
- All Transformer writes go through a `ConcurrentTransformer` mutex wrapper
- Each goroutine gets a `Context.Clone()` with independent evaluation state
- Progress reporting uses atomic counters
- Panics in goroutines are recovered and reported as errors
- Cancellation propagates immediately via `context.WithCancel`

### Streaming + Parallel

These are mutually exclusive. If both are set, parallel takes precedence. For truly massive outputs, use streaming (it wins on both speed and memory).

## Auto-mode: let XLFill decide

Instead of choosing manually, let XLFill analyze your template and pick the optimal mode:

```go
xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithAutoMode(map[string]any{
        "itemCount": len(employees), // hint: how many items
    }),
)
```

**Decision logic:**

| Item count | Streaming-eligible? | Parallel-eligible? | Result |
|---|---|---|---|
| >= 10,000 | Yes | — | **Streaming** |
| >= 100 | — | Yes (multi-core) | **Parallel** (capped at 8 goroutines) |
| >= 1,000 | Yes | No | **Streaming** (fallback) |
| Any | No | No | **Sequential** |

**Streaming blockers:** formulas, images, hyperlinks, multisheet, direction=RIGHT
**Parallel blockers:** nested each/repeat, multisheet, direction=RIGHT, no each commands

### Explicit mode suggestion

For more control, call `SuggestMode` directly and inspect the recommendation:

```go
suggestion, err := xlfill.SuggestMode("template.xlsx", map[string]any{
    "itemCount": 50000,
})
fmt.Println(suggestion.Mode)    // "streaming"
fmt.Println(suggestion.Reasons) // ["large dataset (>=10K items)", "template is streaming-compatible"]

// Apply the suggestion
opts := suggestion.Apply()
xlfill.Fill("template.xlsx", "report.xlsx", data, opts...)
```

## Compiled templates: amortize parsing

When generating the same report with different data (batch jobs, API endpoints, queue workers), parse the template once and reuse:

```go
compiled, err := xlfill.Compile("template.xlsx",
    xlfill.WithRecalculateOnOpen(true),
)

// Fill with different data sets — template bytes cached in memory
for _, dataset := range datasets {
    compiled.Fill(dataset, fmt.Sprintf("report_%d.xlsx", i))
}
```

Each `Fill` call creates a fresh transformer from cached bytes — no file I/O. Options like streaming, parallel, auto-mode, and strict mode are all propagated to each fill.

## Context cancellation and progress

For long-running fills, use Go's standard `context.Context` for cancellation and timeouts:

```go
ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
defer cancel()

err := xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithContext(ctx),
    xlfill.WithProgressFunc(func(p xlfill.FillProgress) {
        fmt.Printf("Processed %d rows on %s\n", p.ProcessedRows, p.CurrentSheet)
    }),
)
```

Progress works with all modes — sequential, streaming, and parallel (using atomic counters for thread safety).

:::caution[`progressFunc` is called concurrently in parallel mode]
With `WithParallelism(n)` where n > 1, your callback fires from multiple goroutines simultaneously. Make sure it's safe for concurrent use — write to a `chan FillProgress`, guard shared state with a mutex, or use atomic counters. The `Elapsed` field is not populated; track wall-clock time yourself if you need it.
:::

## Deferred commands

Several commands use **deferred execution** — they collect their configuration during template processing but apply their effects only after all rows are written. This is both a performance optimization and a correctness requirement: these commands need to know the final output row count to set correct ranges.

**Deferred commands:**

| Command | Why deferred |
|---------|-------------|
| `jx:table` | Table range must cover all output rows |
| `jx:chart` | Chart data ranges must reference final row positions |
| `jx:conditionalFormat` | Format rules must span the entire output range |
| `jx:group` | Outline group ranges depend on final row positions |
| `jx:definedName` | Named ranges must cover all output rows |
| `jx:sparkline` | Data ranges must reference final cell positions |

**How it works:** During template processing, each deferred command records a `DeferredAction` with its template-relative area and attributes. After all rows are written, the engine replays these actions with adjusted row offsets. This means:

- No wasted work if a `jx:if` excludes the area containing the deferred command
- Correct ranges even with nested loops that expand to variable heights
- Compatible with streaming mode (deferred actions run after the stream is finalized)

**Performance impact:** Deferred execution adds negligible overhead — typically < 1ms for a report with multiple tables, charts, and conditional formats. The alternative (applying during expansion and then re-adjusting ranges) would be both slower and more error-prone.

## Internal optimizations

These happen automatically — no configuration needed:

| Optimization | Impact |
|---|---|
| **Differential context map** | Loop variable updates modify the cached map in-place instead of rebuilding. Eliminates ~30K map copies for 10K rows. |
| **Expression compilation cache** | Expressions are compiled once via `sync.Map`. Subsequent evaluations hit the cache (~5M evals/sec). |
| **Pre-allocated slices** | Comment and formula cell lists are pre-sized during template loading. Reduces GC pressure for large templates. |
| **Atomic progress counters** | `Area.rowsProcessed` uses `atomic.Int64` — safe for parallel mode with zero contention. |

## Reproducing the benchmarks

The numbers in this guide come from `go test -bench=. -benchmem` against `benchmark_test.go`. The absolute times depend on your hardware, but the **relative** numbers (streaming vs sequential, sequential vs parallel) should hold on any modern machine.

```bash
# All benchmarks
go test -bench=. -benchmem -run=^$ -timeout=10m

# Just the streaming-vs-sequential comparison
go test -bench=BenchmarkFill -benchmem -run=^$

# CPU profile, to find hot spots in your real template
go test -bench=. -cpuprofile=cpu.prof -run=^$
go tool pprof -http=:8080 cpu.prof
```

Benchmark with **your template**, not the demo one. Hot expressions, deep nesting, and lots of formulas can change which mode wins.

## Tips

1. **Benchmark your actual template** — the examples above use a simple 3-column template. Complex expressions, formulas, and nested loops change the equation.
2. **Streaming is the biggest win** — if your template is compatible, streaming mode gives ~3x speedup and ~60% less memory with zero code changes.
3. **Auto-mode is safe** — it only selects modes your template supports. No silent failures.
4. **Check `Filler.Warnings()` after streaming** — dropped hyperlinks and skipped sheets surface there.
5. **Compile for batch** — if you generate the same report more than once, `Compile` pays for itself on the second fill.
6. **Use `context.Context`** — always set a timeout for server-side report generation to prevent runaway fills.

## What's next?

For raw benchmark numbers and scaling characteristics:

**[Performance Benchmarks &rarr;](/xlfill/reference/performance/)**

For error handling and validation:

**[Error Handling &rarr;](/xlfill/guides/error-handling/)**
