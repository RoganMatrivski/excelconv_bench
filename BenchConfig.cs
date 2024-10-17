using System.Buffers;
using System.Diagnostics.Tracing;

using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Diagnostics.Windows;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Csv;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Environments;
using BenchmarkDotNet.Jobs;
using Microsoft.Diagnostics.NETCore.Client;
using Microsoft.Diagnostics.Tracing.Parsers;

public class BenchConfig
{
    public static IConfig CreateBenchConfig(CommonBenchProperties opts)
    {
        var cfg = ManualConfig.CreateMinimumViable();
        var job = Job.Default;

        if (opts.EtwProfiler) cfg.AddDiagnoser(new EtwProfiler());
        if (opts.EventPipeProfiler)
        {
            var providers = new[]
            {
                new EventPipeProvider(ClrTraceEventParser.ProviderName, EventLevel.Verbose,
                    (long) (ClrTraceEventParser.Keywords.Exception
                    | ClrTraceEventParser.Keywords.GC
                    | ClrTraceEventParser.Keywords.Jit
                    | ClrTraceEventParser.Keywords.JitTracing // for the inlining events
					| ClrTraceEventParser.Keywords.Loader
                    | ClrTraceEventParser.Keywords.NGen)),
                new EventPipeProvider("System.Buffers.ArrayPoolEventSource", EventLevel.Informational, long.MaxValue),
            };

            cfg.AddDiagnoser(new EventPipeProfiler(providers: providers));
        }
        if (opts.MemoryDiagnoser) cfg.AddDiagnoser(MemoryDiagnoser.Default);
        if (opts.ExportRPlot)
        {
            cfg.AddExporter(CsvMeasurementsExporter.Default);
            cfg.AddExporter(RPlotExporter.Default);
        }

        // if (opts.WarmupCount.HasValue)
        // {
        //     job.WithWarmupCount(opts.WarmupCount.Value);
        // }

        // if (opts.IterationCount.HasValue)
        // {
        //     job.WithIterationCount(opts.IterationCount.Value);
        // }

        cfg.AddLogger(ConsoleLogger.Default).AddColumnProvider(DefaultColumnProviders.Instance)
        .AddExporter(CsvExporter.Default)
        .AddExporter(HtmlExporter.Default)
        .AddExporter(PlainExporter.Default);

        return cfg;
    }
}