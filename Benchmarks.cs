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
    [Params(100, 500, 1000)]
    public int ROW;
    [Params(100, 200)]
    public int COL;

    [GlobalSetup]
    public void Setup()
    {
        data = Generator.GenerateDataTable(ROW, COL);
    }


    [Benchmark]
    public void SetWidth_AutoSizeColumn() => new ConverterNPOI(data).SetupSheet().SetWidth_AutoSizeColumn().WriteToStream(Stream.Null);
    [Benchmark]
    public void NPOI_SkiaSharp() => new ConverterNPOI(data).SetupSheet().SetWidth_SkiaSharp().WriteToStream(Stream.Null);
    [Benchmark]
    public void NPOI_CharCount() => new ConverterNPOI(data).SetupSheet().SetWidth_SkiaSharp().WriteToStream(Stream.Null);

    [Benchmark]
    public void ClosedXML_SetWidth_AutoSizeColumn() => new ConverterClosedXML(data).Setup().SetWidth_AutoSizeColumn().WriteToStream(Stream.Null);
    [Benchmark]
    public void ClosedXML_SkiaSharp() => new ConverterClosedXML(data).Setup().SetWidth_SkiaSharp().WriteToStream(Stream.Null);
    [Benchmark]
    public void ClosedXML_CharCount() => new ConverterClosedXML(data).Setup().SetWidth_CharCount().WriteToStream(Stream.Null);

    [Benchmark]
    public void OpenXML_SkiaSharp() => new ConverterOpenXML(Stream.Null, data).SetSetupColumns_SkiaSharp().ExportTable().Dispose();
    [Benchmark]
    public void OpenXML_CharCount() => new ConverterOpenXML(Stream.Null, data).SetSetupColumns_CharCount().ExportTable().Dispose();
}

// [EventPipeProfiler(EventPipeProfile.CpuSampling)]
public class TestSingleBenchmark
{
    private const int COL = 100;
    private const int ROW = 100;
    private DataTable data;

    [GlobalSetup]
    public void Setup()
    {
        data = Generator.GenerateDataTable(ROW, COL);
    }

    [Benchmark]
    public DataTable DatasetGeneration() => Generator.GenerateDataTable(ROW, COL);

    [Benchmark]
    public void NPOI_SetWidth_AutoSizeColumn() => new ConverterNPOI(data).SetupSheet().SetWidth_AutoSizeColumn().WriteToStream(Stream.Null);
    [Benchmark]
    public void NPOI_SkiaSharp() => new ConverterNPOI(data).SetupSheet().SetWidth_SkiaSharp().WriteToStream(Stream.Null);
    [Benchmark]
    public void NPOI_CharCount() => new ConverterNPOI(data).SetupSheet().SetWidth_CharCount().WriteToStream(Stream.Null);

    [Benchmark]
    public void ClosedXML_SetWidth_AutoSizeColumn() => new ConverterClosedXML(data).Setup().SetWidth_AutoSizeColumn().WriteToStream(Stream.Null);
    [Benchmark]
    public void ClosedXML_SkiaSharp() => new ConverterClosedXML(data).Setup().SetWidth_SkiaSharp().WriteToStream(Stream.Null);
    [Benchmark]
    public void ClosedXML_CharCount() => new ConverterClosedXML(data).Setup().SetWidth_CharCount().WriteToStream(Stream.Null);

    [Benchmark]
    public void OpenXML_SkiaSharp() => new ConverterOpenXML(Stream.Null, data).SetSetupColumns_SkiaSharp().ExportTable().Dispose();
    [Benchmark]
    public void OpenXML_CharCount() => new ConverterOpenXML(Stream.Null, data).SetSetupColumns_CharCount().ExportTable().Dispose();
}


[EventPipeProfiler(EventPipeProfile.CpuSampling)]
public class TestSingleBenchmark_Dynamic
{
    private int ROW;
    private int COL;

    // [Global]
}
