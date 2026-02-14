---
title: Performance
description: XLFill benchmarks, scaling characteristics, and optimization tips.
---

XLFill is designed to be fast enough that performance is never a concern for typical report generation. Here are the numbers.

## Benchmarks

Measured on Intel i5-9300H @ 2.40GHz:

| Scenario | Rows | Time | Memory | Throughput |
|----------|------|------|--------|------------|
| Simple 3-column template | 100 | 6.2ms | 2.0 MB | ~16,000 rows/sec |
| Simple 3-column template | 1,000 | 35ms | 9.6 MB | ~28,500 rows/sec |
| Simple 3-column template | 10,000 | 308ms | 86.8 MB | ~32,500 rows/sec |
| Nested loops (10 groups x 20 items) | 200 | 2.6ms | 920 KB | ~77,000 rows/sec |
| Single expression evaluation | 1 | 199ns | 48 B | ~5M evals/sec |
| Single comment parse | 1 | 4.1us | 1 KB | ~244K parses/sec |

## Key characteristics

**Linear scaling** — processing time grows linearly with the number of rows. 10x more rows = ~10x more time. No surprises.

**~8.7 KB per row** at scale — memory usage is dominated by the Excel file representation in excelize. For 10,000 rows, expect ~87 MB.

**Expression caching** — expressions are compiled once on first encounter and cached. Subsequent evaluations of the same expression hit the cache, giving ~5 million evaluations per second.

**Comment parsing** — template comments are parsed at ~244K per second. Even a template with hundreds of commands parses in under a millisecond.

## What this means in practice

| Report size | Expected time |
|-------------|--------------|
| Small (100 rows) | Instant — under 10ms |
| Medium (1,000 rows) | ~35ms — imperceptible |
| Large (10,000 rows) | ~300ms — still fast |
| Very large (100,000 rows) | ~3 seconds — consider streaming |

For comparison, the time to open a 10,000-row Excel file in the Excel application is typically longer than the time XLFill takes to generate it.

## Tips for large reports

1. **Keep templates simple** — fewer expressions per row means faster evaluation
2. **Filter early with `select`** — `jx:each(... select="e.Active")` is faster than generating rows you'll discard
3. **Avoid deep nesting** — each nesting level multiplies the work. Three levels deep with large collections can add up
4. **Minimize formula references** — formula expansion scans all formulas after each command. Templates with hundreds of formulas in the expanded area may see overhead

## Requirements

- Go 1.24+
- `.xlsx` files only (the `.xls` binary format is not supported)

---

You've reached the end of the documentation. If you haven't already, start with the [Getting Started](/xlfill/guides/getting-started/) guide to build your first template-driven report, or check out [Why XLFill?](/xlfill/guides/why-xlfill/) to see the full case for the template-first approach.
