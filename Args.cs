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
    public class GenerateOptions
    {
        [Option('r', "rows", Default = 1000000, HelpText = "Number of rows (max 5,000,000)")]
        public int Rows
        {
            get => _rows;
            set => _rows = Math.Min(value, 5000000);
        }
        private int _rows;

        [Option('c', "columns", Default = 256, HelpText = "Number of columns (max 512)")]
        public int Columns
        {
            get => _columns;
            set => _columns = Math.Min(value, 512);
        }
        private int _columns;

        [Value(0, MetaName = "path", Required = true, HelpText = "Output file path")]
        public string Path { get; set; }
    }


    [Verb("bench", HelpText = "Benchmark")]
    public class BenchmarkOptions { }
    [Verb("singleiterbench", HelpText = "Benchmark")]
    public class SingleBenchmarkOptions { }

    [Verb("trace-npoi", HelpText = "Benchmark")]
    public class TraceNPOI { }
    [Verb("trace-openxml", HelpText = "Benchmark")]
    public class TraceOpenXML { }
}
