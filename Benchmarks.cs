using System.Data;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.Windows.Configs;
using BenchmarkDotNet.Diagnosers;

// [EtwProfiler(performExtraBenchmarksRun: false)]
// [EventPipeProfiler(EventPipeProfile.CpuSampling)]
public class TestBenchmark
{
    // private const int COL = 256;
    private DataTable data;

    // [Params(100, 200, 500, 1000)]
    // [Params(8, 16, 32, 64, 128)]
    [Params(5, 10, 15, 20, 30, 40, 60, 80, 100)]
    public int ROW;
    [Params(5, 10, 15, 20, 30, 40, 60, 80, 100)]
    public int COL;

    [GlobalSetup]
    public void Setup()
    {
        data = Generator.GenerateDataTable(ROW, COL);
    }

    [Benchmark]
    public byte[] NPOI() => OldConverter.ConvertToExcel(data, "1234");

    [Benchmark]
    public byte[] OpenXML() => AppParser.ConvertTable(data);
}

[EventPipeProfiler(EventPipeProfile.CpuSampling)]
public class TestSingleBenchmark
{
    private const int COL = 100;
    private const int ROW = 1_000_000;
    private DataTable data;

    // [Params(100, 200, 500, 1000)]
    // [Params(8, 16, 32, 64, 128)]
    // [Params(5, 10, 15, 20, 30, 40, 60, 80, 100)]
    // public int ROW;
    // [Params(5, 10, 15, 20, 30, 40, 60, 80, 100)]
    // public int COL;

    [GlobalSetup]
    public void Setup()
    {
        data = Generator.GenerateDataTable(ROW, COL);
    }

    [Benchmark]
    public byte[] NPOI() => OldConverter.ConvertToExcel(data, "1234");
    [Benchmark]
    public byte[] NPOI_new() => OldConverter.ConvertToExcel_New(data, "1234");

    [Benchmark]
    public byte[] OpenXML() => AppParser.ConvertTable(data);
}
