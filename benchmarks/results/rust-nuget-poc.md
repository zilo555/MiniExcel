# Rust-backed NuGet proof of concept

Date: 2026-08-30

## Scope

This proof of concept adds `MiniExcelRust.Query` as a synchronous, path-based dynamic XLSX
query. A managed `MiniExcel.Rust` package calls a Rust `cdylib` through a versioned C ABI and
transfers up to 64 rows per P/Invoke call in a tagged binary frame.

The test consumer references `MiniExcel.Rust` only through a freshly packed local NuGet feed.
The managed comparison references the current MiniExcel V2 source tree.

## Environment

| Item | Value |
| --- | --- |
| Operating system | Windows x64 |
| .NET SDK | 10.0.103 |
| Rust | 1.85.0 |
| Workbook | `Test100,000x10.xlsx` |
| Workbook rows and cells | 100,000 rows, 1,000,000 cells per pass |
| Local package | `MiniExcel.Rust.0.1.0-preview.1.nupkg` |
| Local package size | 639,093 bytes (0.61 MB) |
| Native asset | `runtimes/win-x64/native/miniexcel_ffi.dll` |

## Correctness and lifecycle

The local package consumer compared column order and every value against MiniExcel V2 for:

- Headerless dynamic query.
- Header-based dynamic query.
- Empty rows with and without a header.

All 15 compared rows matched. A separate lifecycle check opened 100 queries, consumed one row,
disposed each iterator early, and then deleted the workbook copy successfully.

## Cold query

Five fresh processes performed one measured pass without an in-process warm-up. Values are
averages.

| Runtime | Elapsed | First row | Managed allocation | Peak working set |
| --- | ---: | ---: | ---: | ---: |
| Managed V2 | 2,530.10 ms | 171.11 ms | 668.93 MB | 67.53 MB |
| Rust NuGet | 725.78 ms | 35.76 ms | 142.49 MB | 45.38 MB |

The Rust NuGet path was 3.49 times faster, reduced first-row latency by 79.1%, reduced managed
allocation by 78.7%, and reduced peak working set by 32.8%.

## Warm sustained query

Five fresh processes performed one untimed warm-up pass followed by three measured passes. Values
are averages for all three measured passes.

| Runtime | Elapsed | First measured row | Managed allocation | Peak working set |
| --- | ---: | ---: | ---: | ---: |
| Managed V2 | 5,222.79 ms | 44.97 ms | 2,006.98 MB | 78.25 MB |
| Rust NuGet | 1,863.37 ms | 28.51 ms | 427.46 MB | 50.49 MB |

The Rust NuGet path was 2.80 times faster, reduced first-row latency by 36.6%, reduced managed
allocation by 78.7%, and reduced peak working set by 35.5%.

Managed allocation excludes allocations inside the Rust native library. Peak working set includes
the complete process and is the more useful cross-runtime memory measurement. Similar peak working
sets across one and repeated passes show no pass-count-proportional retention in this test; broader
workbook-size scaling is still required for a general bounded-memory claim.

## Acceptance status

| Requirement | Status |
| --- | --- |
| In-process throughput at least current V2 | Pass |
| First-row latency no more than 10% slower | Pass; it is faster |
| Focused dynamic row equivalence | Pass |
| Deterministic early disposal | Pass |
| No repeated-pass memory growth | Pass for this workbook |
| Complete dynamic parity contract | Not yet tested |
| Cancellation latency | Not implemented in this synchronous API |
| All supported RIDs | Not met; package currently contains only `win-x64` |
| Single-file, trimming, and NativeAOT | Not yet tested |

## Reproduce

From the MiniExcel repository root:

```powershell
pwsh ./benchmarks/test-rust-nuget.ps1 -Iterations 5 -Passes 3 -WarmupPasses 1
pwsh ./benchmarks/test-rust-nuget.ps1 -Iterations 5 -Passes 1 -WarmupPasses 0
```

The result supports continuing with the optional Rust query backend. It does not yet support
replacing the main MiniExcel NuGet implementation or enabling Rust by default.