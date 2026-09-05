using System.Diagnostics;
using System.Text.Json;
using MiniExcelLib;
using MiniExcelLib.OpenXml;
using RustQuery = MiniExcelLibs.MiniExcelRust;

if (args.Length < 1)
    return Usage();

return args[0].ToLowerInvariant() switch
{
    "verify" => Verify(args),
    "lifecycle" => Lifecycle(args),
    "managed" => Benchmark(args, useRust: false),
    "rust" => Benchmark(args, useRust: true),
    _ => Usage()
};

static int Verify(string[] args)
{
    if (args.Length != 3 || !bool.TryParse(args[2], out var useHeaderRow))
        return Usage();

    var path = Path.GetFullPath(args[1]);
    using var managed = QueryManaged(path, useHeaderRow).GetEnumerator();
    using var rust = RustQuery.Query(path, useHeaderRow).GetEnumerator();
    var rowIndex = 0;

    while (true)
    {
        var hasManaged = managed.MoveNext();
        var hasRust = rust.MoveNext();
        if (hasManaged != hasRust)
            throw new InvalidOperationException($"Row count differs after row {rowIndex}.");
        if (!hasManaged)
            break;

        CompareRow(managed.Current, rust.Current, rowIndex);
        rowIndex++;
    }

    Console.WriteLine($"Verified {rowIndex} rows from the local MiniExcel.Rust NuGet package.");
    return 0;
}

static int Lifecycle(string[] args)
{
    if (args.Length != 2)
        return Usage();

    var temporaryPath = Path.Combine(Path.GetTempPath(), $"miniexcel-rust-{Guid.NewGuid():N}.xlsx");
    File.Copy(Path.GetFullPath(args[1]), temporaryPath);
    try
    {
        for (var iteration = 0; iteration < 100; iteration++)
        {
            using var rows = RustQuery.Query(temporaryPath).GetEnumerator();
            if (!rows.MoveNext())
                throw new InvalidOperationException("The lifecycle fixture did not contain a row.");
        }

        File.Delete(temporaryPath);
        Console.WriteLine("Verified 100 early-disposal query cycles from the local MiniExcel.Rust NuGet package.");
        return 0;
    }
    finally
    {
        if (File.Exists(temporaryPath))
            File.Delete(temporaryPath);
    }
}

static int Benchmark(string[] args, bool useRust)
{
    if (args.Length is < 2 or > 4)
        return Usage();

    var path = Path.GetFullPath(args[1]);
    var passes = args.Length >= 3 ? int.Parse(args[2]) : 1;
    var warmupPasses = args.Length >= 4 ? int.Parse(args[3]) : 0;
    if (passes < 1 || warmupPasses < 0)
        throw new ArgumentOutOfRangeException(nameof(args), "Passes must be positive and warm-up passes cannot be negative.");

    for (var pass = 0; pass < warmupPasses; pass++)
        Consume(path, useRust);

    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();

    var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
    var stopwatch = Stopwatch.StartNew();
    var firstRowMilliseconds = 0d;
    long rowCount = 0;
    long cellCount = 0;

    for (var pass = 0; pass < passes; pass++)
    {
        var rows = useRust ? RustQuery.Query(path) : QueryManaged(path, useHeaderRow: false);
        foreach (var row in rows)
        {
            if (rowCount == 0)
                firstRowMilliseconds = stopwatch.Elapsed.TotalMilliseconds;
            rowCount++;
            cellCount += row.Count;
        }
    }

    stopwatch.Stop();
    var result = new BenchmarkResult(
        useRust ? "RustNuGet" : "ManagedV2",
        passes,
        rowCount,
        cellCount,
        stopwatch.Elapsed.TotalMilliseconds,
        firstRowMilliseconds,
        GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore);
    Console.WriteLine(JsonSerializer.Serialize(result));
    return 0;
}

static void Consume(string path, bool useRust)
{
    var rows = useRust ? RustQuery.Query(path) : QueryManaged(path, useHeaderRow: false);
    foreach (var row in rows)
        _ = row.Count;
}

static IEnumerable<IDictionary<string, object?>> QueryManaged(string path, bool useHeaderRow)
{
    var importer = MiniExcel.Importers.GetOpenXmlImporter();
    foreach (IDictionary<string, object?> row in importer.Query(path, hasHeaderRow: useHeaderRow))
        yield return row;
}

static void CompareRow(
    IDictionary<string, object?> expected,
    IDictionary<string, object?> actual,
    int rowIndex)
{
    if (!expected.Keys.SequenceEqual(actual.Keys, StringComparer.Ordinal))
        throw new InvalidOperationException($"Column order differs at row {rowIndex}.");

    foreach (var key in expected.Keys)
    {
        if (!Equals(expected[key], actual[key]))
            throw new InvalidOperationException(
                $"Value differs at row {rowIndex}, column {key}: managed={expected[key] ?? "<null>"}, rust={actual[key] ?? "<null>"}.");
    }
}

static int Usage()
{
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  MiniExcel.Rust.PackageTests verify <xlsx-path> <use-header-row>");
    Console.Error.WriteLine("  MiniExcel.Rust.PackageTests lifecycle <xlsx-path>");
    Console.Error.WriteLine("  MiniExcel.Rust.PackageTests <managed|rust> <xlsx-path> [passes] [warmup-passes]");
    return 2;
}

internal sealed record BenchmarkResult(
    string Runtime,
    int Passes,
    long Rows,
    long Cells,
    double ElapsedMilliseconds,
    double FirstRowMilliseconds,
    long AllocatedBytes);