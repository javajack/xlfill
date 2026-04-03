---
title: Performance
description: XLFill benchmarks, scaling characteristics, processing modes, and optimization tips.
---

XLFill is designed to be fast enough that performance is never a concern for typical report generation. For large workloads, streaming and parallel modes push throughput even further.

## Benchmarks

Measured on Intel i5-9300H @ 2.40GHz:

### Sequential mode (default)

| Scenario | Rows | Time | Memory | Throughput |
|----------|------|------|--------|------------|
| Simple 3-column template | 100 | 4.9ms | 1.7 MB | ~20,000 rows/sec |
| Simple 3-column template | 1,000 | 27ms | 8.3 MB | ~37,000 rows/sec |
| Simple 3-column template | 10,000 | 250ms | 75 MB | ~40,000 rows/sec |
| Nested loops (10 groups x 20 items) | 200 | 2.0ms | 810 KB | ~100,000 rows/sec |
| Single expression evaluation | 1 | 182ns | 48 B | ~5.5M evals/sec |

### Streaming mode (`WithStreaming(true)`)

| Scenario | Rows | Time | Memory | vs Sequential |
|----------|------|------|--------|---------------|
| Simple 3-column template | 1,000 | **8.9ms** | **3.3 MB** | **3x faster, 60% less memory** |

### Parallel mode (`WithParallelism(n)`)

| Scenario | Rows | Goroutines | Time | vs Sequential |
|----------|------|------------|------|---------------|
| Simple 3-column template | 1,000 | 2 | 38ms | Overhead-dominated at this scale |
| Simple 3-column template | 1,000 | 4 | 37ms | Overhead-dominated at this scale |

Parallel mode shines with CPU-bound expression evaluation (complex expressions, many columns). For I/O-dominated workloads like the simple benchmark above, sequential is faster due to zero mutex overhead.

## Key characteristics

**Linear scaling** — processing time grows linearly with the number of rows. 10x more rows = ~10x more time. No surprises.

**~7.5 KB per row** at scale in sequential mode — memory is dominated by the Excel file representation in excelize. Streaming mode cuts this to ~3.3 KB per row.

**Expression caching** — expressions are compiled once on first encounter and cached via `sync.Map`. Subsequent evaluations of the same expression hit the cache, giving ~5.5 million evaluations per second.

**Differential context updates** — loop variable changes use in-place map updates instead of full rebuilds. This eliminates ~30,000 map copies for a 10K-row loop — a 19% end-to-end speedup compared to the naive approach.

**Pre-allocated slices** — comment and formula cell lists are pre-sized during template loading, reducing GC pressure for large templates.

## What this means in practice

| Report size | Sequential | Streaming | Recommendation |
|-------------|-----------|-----------|----------------|
| Small (100 rows) | < 5ms | — | Sequential (default) |
| Medium (1,000 rows) | ~27ms | ~9ms | Sequential or streaming |
| Large (10,000 rows) | ~250ms | ~80ms* | Streaming if compatible |
| Very large (100,000 rows) | ~2.5s | ~0.8s* | Streaming strongly recommended |

*Streaming estimates based on measured 3x speedup ratio.

Or let XLFill decide: `WithAutoMode(map[string]any{"itemCount": len(rows)})` analyzes your template and picks the optimal mode automatically. See the [Performance Tuning guide](/xlfill/guides/performance-tuning/).

## Tips for large reports

1. **Try streaming first** — if your template has no formulas, images, or hyperlinks, `WithStreaming(true)` gives the biggest win with zero code changes
2. **Use `WithAutoMode`** — it picks the right mode based on your template structure and data size
3. **Compile for batch** — `xlfill.Compile("template.xlsx")` caches the template bytes; each fill avoids file I/O
4. **Set a context timeout** — `WithContext(ctx)` prevents runaway fills on server-side endpoints
5. **Keep templates simple** — fewer expressions per row means faster evaluation
6. **Filter early with `select`** — `jx:each(... select="e.Active")` is faster than generating rows you'll discard
7. **Avoid deep nesting** — each nesting level multiplies the work. Three levels deep with large collections can add up

## Requirements

- Go 1.24+
- `.xlsx` files only (the `.xls` binary format is not supported)

## What's next?

For detailed streaming, parallel, and compiled template configuration:

**[Performance Tuning &rarr;](/xlfill/guides/performance-tuning/)**

Having trouble with a template? XLFill ships with built-in tools for inspecting and validating templates:

**[Debugging & Troubleshooting &rarr;](/xlfill/guides/debugging/)**
