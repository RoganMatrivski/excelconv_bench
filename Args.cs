using CommandLine;
namespace Args
{
    [Verb("convert", HelpText = "Convert files")]
    public class ConvertOptions
    {

        [Value(0, MetaName = "path", Required = true, HelpText = "Input file path")]
        public string Path { get; set; }

        [Value(1, MetaName = "outputPath", Required = true, HelpText = "Output file path")]
        public string OutputPath { get; set; }
    }


    [Verb("generate", HelpText = "Generate files")]
    public class GenerateOptions : CommonGenerateArgs
    {
        [Value(0, MetaName = "path", Required = true, HelpText = "Output file path")]
        public string Path { get; set; }
    }


    [Verb("bench", HelpText = "Benchmark")]
    public class BenchmarkOptions : CommonBenchProperties { }
    [Verb("singleiterbench", HelpText = "Benchmark")]
    public class SingleBenchmarkOptions : CommonBenchProperties { }

    [Verb("trace-npoi", HelpText = "Benchmark")]
    public class TraceNPOI : CommonGenerateArgs { }
    [Verb("trace-openxml", HelpText = "Benchmark")]
    public class TraceOpenXML : CommonGenerateArgs { }
    [Verb("trace-closedxml", HelpText = "Benchmark")]
    public class TraceClosedXML : CommonGenerateArgs { }
    [Verb("singlerunall")]
    public class SingleRunAll : CommonGenerateArgs { }
}

public class CommonBenchProperties
{
    [Option("etw-profiler", HelpText = "Enable EtwProfiler")]
    public bool EtwProfiler { get; set; }
    [Option("event-pipe-profiler", HelpText = "Enable EventPipeProfiler")]
    public bool EventPipeProfiler { get; set; }
    [Option("memory-diagnoser", HelpText = "Enable MemoryDiagnoser")]
    public bool MemoryDiagnoser { get; set; }
    [Option("export-rplot", HelpText = "Enable R Plot exporters")]
    public bool ExportRPlot { get; set; }
}

public class CommonGenerateArgs {
	[Option('r', "rows", Default = 10000, HelpText = "Number of rows (max 5,000,000)")]
	public int Rows
	{
		get => _rows;
		set => _rows = Math.Min(value, 5000000);
	}
	private int _rows;

	[Option('c', "columns", Default = 100, HelpText = "Number of columns (max 512)")]
	public int Columns
	{
		get => _columns;
		set => _columns = Math.Min(value, 512);
	}
	private int _columns;
}