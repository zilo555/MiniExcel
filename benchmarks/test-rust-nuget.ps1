param(
    [ValidateRange(1, 20)]
    [int]$Iterations = 5,
    [ValidateRange(1, 20)]
    [int]$Passes = 3,
    [ValidateRange(0, 20)]
    [int]$WarmupPasses = 1
)

$ErrorActionPreference = "Stop"
$repositoryRoot = Split-Path $PSScriptRoot -Parent
$rustRepository = Join-Path (Split-Path $repositoryRoot -Parent) "MiniExcel-Rust"
$localFeed = Join-Path $repositoryRoot "artifacts/local-rust"
$localPackages = Join-Path $localFeed "packages"
$packageProject = Join-Path $repositoryRoot "src/MiniExcel.Rust/MiniExcel.Rust.csproj"
$consumerProject = Join-Path $PSScriptRoot "MiniExcel.Rust.PackageTests/MiniExcel.Rust.PackageTests.csproj"
$consumerDll = Join-Path $PSScriptRoot "MiniExcel.Rust.PackageTests/bin/Release/net10.0/MiniExcel.Rust.PackageTests.dll"
$benchmarkWorkbook = Join-Path $PSScriptRoot "MiniExcel.Benchmarks/data/Test100,000x10.xlsx"
$fixtures = @(
    @{ Path = "tests/data/xlsx/TestDynamicQueryBasic_WithoutHead.xlsx"; Header = "false" },
    @{ Path = "tests/data/xlsx/TestDynamicQueryBasic.xlsx"; Header = "true" },
    @{ Path = "tests/data/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx"; Header = "false" },
    @{ Path = "tests/data/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx"; Header = "true" }
)

New-Item -ItemType Directory -Force $localFeed | Out-Null
Remove-Item $localPackages -Recurse -Force -ErrorAction SilentlyContinue

Push-Location $rustRepository
try {
    & cargo +1.85.0 build --release -p miniexcel-ffi --locked
    if ($LASTEXITCODE -ne 0) { throw "Rust FFI release build failed." }
}
finally {
    Pop-Location
}

& dotnet pack $packageProject -c Release -o $localFeed --nologo --no-restore
if ($LASTEXITCODE -ne 0) { throw "MiniExcel.Rust pack failed." }

& dotnet restore $consumerProject --packages $localPackages --force --no-cache --nologo
if ($LASTEXITCODE -ne 0) { throw "Local package consumer restore failed." }
& dotnet build $consumerProject -c Release --no-restore --nologo
if ($LASTEXITCODE -ne 0) { throw "Local package consumer build failed." }

foreach ($fixture in $fixtures) {
    & dotnet $consumerDll verify (Join-Path $repositoryRoot $fixture.Path) $fixture.Header
    if ($LASTEXITCODE -ne 0) { throw "Local package verification failed for $($fixture.Path)." }
}
& dotnet $consumerDll lifecycle (Join-Path $repositoryRoot $fixtures[0].Path)
if ($LASTEXITCODE -ne 0) { throw "Local package lifecycle verification failed." }

function Invoke-BenchmarkProcess {
    param([string]$Runtime, [int]$Iteration)

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = "dotnet"
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @($consumerDll, $Runtime, $benchmarkWorkbook, "$Passes", "$WarmupPasses")) {
        $startInfo.ArgumentList.Add($argument)
    }

    $process = [Diagnostics.Process]::Start($startInfo)
    $peakWorkingSet = 0L
    while (-not $process.WaitForExit(10)) {
        $process.Refresh()
        $peakWorkingSet = [Math]::Max($peakWorkingSet, $process.WorkingSet64)
    }

    $output = $process.StandardOutput.ReadToEnd().Trim()
    $errorOutput = $process.StandardError.ReadToEnd().Trim()
    if ($process.ExitCode -ne 0) {
        throw "$Runtime benchmark failed: $errorOutput"
    }

    $measurement = $output | ConvertFrom-Json
    [pscustomobject]@{
        Runtime = $measurement.Runtime
        Iteration = $Iteration
        Rows = $measurement.Rows
        Cells = $measurement.Cells
        ElapsedMs = [Math]::Round($measurement.ElapsedMilliseconds, 2)
        FirstRowMs = [Math]::Round($measurement.FirstRowMilliseconds, 2)
        AllocatedMB = [Math]::Round($measurement.AllocatedBytes / 1MB, 2)
        PeakWorkingSetMB = [Math]::Round($peakWorkingSet / 1MB, 2)
    }
}

$results = foreach ($iteration in 1..$Iterations) {
    Invoke-BenchmarkProcess -Runtime managed -Iteration $iteration
    Invoke-BenchmarkProcess -Runtime rust -Iteration $iteration
}

if (($results.Rows | Select-Object -Unique).Count -ne 1 -or ($results.Cells | Select-Object -Unique).Count -ne 1) {
    throw "Managed and Rust runners returned different row or cell counts."
}

$results | Format-Table -AutoSize

"Summary (averages)"
$results | Group-Object Runtime | ForEach-Object {
    [pscustomobject]@{
        Runtime = $_.Name
        ElapsedMs = [Math]::Round(($_.Group.ElapsedMs | Measure-Object -Average).Average, 2)
        FirstRowMs = [Math]::Round(($_.Group.FirstRowMs | Measure-Object -Average).Average, 2)
        AllocatedMB = [Math]::Round(($_.Group.AllocatedMB | Measure-Object -Average).Average, 2)
        PeakWorkingSetMB = [Math]::Round(($_.Group.PeakWorkingSetMB | Measure-Object -Average).Average, 2)
    }
} | Format-Table -AutoSize