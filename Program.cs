using CommandLine;
using Args;
using System.Reflection;
using System.Diagnostics;
using CommandLine.Text;
using BenchmarkDotNet.Running;
using System.Threading;
using System.Threading.Tasks;

public class App
{
    static void Main(string[] args)
    {
        var types = LoadVerbs();

        Parser.Default.ParseArguments(args, types)
        .WithParsed(Run);
    }

    private static Type[] LoadVerbs() =>
        Assembly
            .GetExecutingAssembly()
            .GetTypes()
            .Where(t => t.GetCustomAttribute<VerbAttribute>() != null)
            .ToArray();

    private static void Run(object obj)
    {
        var cts = new CancellationTokenSource();
        Task measureMemoryTask = Task.Run(() => MeasureGCMemoryIndefinitely(cts.Token));
		
        switch (obj)
        {
            case GenerateOptions opts:
                Generator.RunGenerateAndReturnExitCode(opts);
                break;
            case BenchmarkOptions opts:
                BenchmarkDotNet.Running.BenchmarkRunner.Run<TestBenchmark>();
                break;
            case SingleBenchmarkOptions opts:
                BenchmarkDotNet.Running.BenchmarkRunner.Run<TestSingleBenchmark>(BenchConfig.CreateBenchConfig(opts));
                break;
            case SingleRunAll opts:
                var dt = Generator.GenerateDataTable(opts.Rows, opts.Columns);
                Console.WriteLine("Done generating table");

                new ConverterNPOI(dt).SetupSheet().SetWidth_AutoSizeColumn().WriteToStream(Stream.Null);
                Console.WriteLine("DONE: SetWidth_AutoSizeColumn");
                new ConverterNPOI(dt).SetupSheet().SetWidth_SkiaSharp().WriteToStream(Stream.Null);
                Console.WriteLine("DONE: SetWidth_SkiaSharp");
                new ConverterNPOI(dt).SetupSheet().SetWidth_CharCount().WriteToStream(Stream.Null);
                Console.WriteLine("DONE: SetWidth_CharCount");
                new ConverterOpenXML(Stream.Null, dt).SetSetupColumns_SkiaSharp().ExportTable();
                Console.WriteLine("DONE: SetSetupColumns_SkiaSharp");
                new ConverterOpenXML(Stream.Null, dt).SetSetupColumns_CharCount().ExportTable();
                Console.WriteLine("DONE: SetSetupColumns_CharCount");

                break;
            case TraceNPOI opts:
                MeasureExecTime(() =>
                {
                    new ConverterNPOI(opts.Rows, opts.Columns)
                        .SetupSheet()
                        .SetWidth_SkiaSharp()
                        .WriteToStream(File.Create("result/npoi.xlsx"));
                }).PrintReadableTime(prepend: "Export time: ");
                break;
            case TraceOpenXML opts:
                MeasureExecTime(() =>
                {
                    using var fs = File.Create("result/openxml.xlsx");
                    new ConverterOpenXML(fs, opts.Rows, opts.Columns)
                        .SetSetupColumns_CharCount()
                        .ExportTable();
                }).PrintReadableTime(prepend: "Export time: ");
                break;
            case TraceClosedXML opts:
                MeasureExecTime(() =>
                {
                    using var fs = File.Create("result/closedxml.xlsx");
                    using var conv = new ConverterClosedXML(opts.Rows, opts.Columns);
                    conv.Setup()
                        .SetWidth_CharCount()
                        .WriteToStream(fs);
                }).PrintReadableTime(prepend: "Export time: ");
                break;
        }
		
        cts.Cancel();
        measureMemoryTask.Wait();
    }

    private static TimeSpan MeasureExecTime(Action fn)
    {
        Stopwatch stopwatch = new();
        stopwatch.Start();
        fn();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }
	
	static void MeasureGCMemoryIndefinitely(CancellationToken token)
    {
        try
        {
            while (!token.IsCancellationRequested)
            {
                // Measure GC total memory in bytes
                long totalMemory = GC.GetTotalMemory(forceFullCollection: false);

                // Convert memory to megabytes
                double totalMemoryMB = totalMemory / (1024.0 * 1024.0);

                // Move the cursor to the beginning of the current line
                Console.SetCursorPosition(0, Console.CursorTop);

                // Overwrite the line with the new memory value
                Console.Write($"GC Total Memory: {totalMemoryMB:F2} MB        ");

                // Wait for a second before measuring again
                Task.Delay(1000).Wait();
            }
        }
        catch (OperationCanceledException)
        {
            // Task was canceled
            Console.WriteLine("\nMemory measurement task canceled.");
        }
    }

}