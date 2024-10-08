using CommandLine;
using Args;

Console.WriteLine("Hello World!");

return Parser.Default.ParseArguments<ConvertOptions, GenerateOptions, BenchmarkOptions, TraceNPOI, TraceOpenXML, SingleBenchmarkOptions>(args)
    .MapResult(
        (ConvertOptions opts) => AppParser.RunConvertAndReturnExitCode(opts),
        (GenerateOptions opts) => Generator.RunGenerateAndReturnExitCode(opts),
        (BenchmarkOptions opts) =>
        {
            BenchmarkDotNet.Running.BenchmarkRunner.Run<TestBenchmark>();
            return 0;
        },
        (TraceNPOI opts) =>
        {
            var asdf = new TestSingleBenchmark();
            asdf.Setup();
            // asdf.NPOI();

            using (var fs = File.Create("result/npoi.xlsx"))
            {
                var filearr = asdf.NPOI();
                fs.Write(filearr);
            }

            return 0;
        },
        (TraceOpenXML opts) =>
        {
            var asdf = new TestSingleBenchmark();
            asdf.Setup();
            // asdf.OpenXML();

            using (var fs = File.Create("result/openxml.xlsx"))
            {
                var filearr = asdf.OpenXML();
                fs.Write(filearr);
            }
            return 0;
        },
        (SingleBenchmarkOptions opts) =>
        {
            BenchmarkDotNet.Running.BenchmarkRunner.Run<TestSingleBenchmark>();
            return 0;
        },
        errs => 1);
