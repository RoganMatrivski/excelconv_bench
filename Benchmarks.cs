using System.Data;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.Windows.Configs;
using BenchmarkDotNet.Diagnosers;

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
    public void NPOI() => OldConverter.ConvertToExcel(data, "1234");

    [Benchmark]
    public void OpenXML() => AppParser.ConvertTable(data);
}

[EtwProfiler(performExtraBenchmarksRun: false)]
[EventPipeProfiler(EventPipeProfile.CpuSampling)]
public class TestSingleBenchmark
{
    private const int COL = 32;
    private const int ROW = 32;
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

    // [Benchmark]
    public byte[] OpenXML() => AppParser.ConvertTable(data);
}
