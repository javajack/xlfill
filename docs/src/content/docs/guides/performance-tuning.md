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

**Limitations:**
- Formula reference remapping is skipped (formulas are written verbatim)
- Hyperlinks are silently written as plain text
- Images are not supported (returns error)
- Single-sheet output only
- Rows must be written in ascending order (guaranteed by template processing)

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

## Internal optimizations

These happen automatically — no configuration needed:

| Optimization | Impact |
|---|---|
| **Differential context map** | Loop variable updates modify the cached map in-place instead of rebuilding. Eliminates ~30K map copies for 10K rows. |
| **Expression compilation cache** | Expressions are compiled once via `sync.Map`. Subsequent evaluations hit the cache (~5M evals/sec). |
| **Pre-allocated slices** | Comment and formula cell lists are pre-sized during template loading. Reduces GC pressure for large templates. |
| **Atomic progress counters** | `Area.rowsProcessed` uses `atomic.Int64` — safe for parallel mode with zero contention. |

## Tips

1. **Benchmark your actual template** — the examples above use a simple 3-column template. Complex expressions, formulas, and nested loops change the equation.
2. **Streaming is the biggest win** — if your template is compatible, streaming mode gives 3x speedup and 60% less memory with zero code changes.
3. **Auto-mode is safe** — it only selects modes your template supports. No silent failures.
4. **Compile for batch** — if you generate the same report more than once, `Compile` pays for itself on the second fill.
5. **Use `context.Context`** — always set a timeout for server-side report generation to prevent runaway fills.

## What's next?

For raw benchmark numbers and scaling characteristics:

**[Performance Benchmarks &rarr;](/xlfill/reference/performance/)**

For error handling and validation:

**[Error Handling &rarr;](/xlfill/guides/error-handling/)**
