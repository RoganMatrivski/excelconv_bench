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

            var watch = System.Diagnostics.Stopwatch.StartNew();

            using (var fs = File.Create("result/npoi.xlsx"))
            {
                var filearr = asdf.NPOI_new();
                fs.Write(filearr);
            }

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            TimeSpan timeSpan = TimeSpan.FromMilliseconds(elapsedMs);
            string readableTime = string.Format("{0:D2}:{1:D2}:{2:D2}.{3:D3}",
                timeSpan.Hours,
                timeSpan.Minutes,
                timeSpan.Seconds,
                timeSpan.Milliseconds);
            Console.WriteLine($"Export time: {readableTime}");

            return 0;
        },
        (TraceOpenXML opts) =>
        {
            var asdf = new TestSingleBenchmark();
            asdf.Setup();
            // asdf.OpenXML();

            var watch = System.Diagnostics.Stopwatch.StartNew();

            using (var fs = File.Create("result/openxml.xlsx"))
            {
                var filearr = asdf.OpenXML();
                fs.Write(filearr);
            }

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            TimeSpan timeSpan = TimeSpan.FromMilliseconds(elapsedMs);
            string readableTime = string.Format("{0:D2}:{1:D2}:{2:D2}.{3:D3}",
                timeSpan.Hours,
                timeSpan.Minutes,
                timeSpan.Seconds,
                timeSpan.Milliseconds);
            Console.WriteLine($"Export time: {readableTime}");

            return 0;
        },
        (SingleBenchmarkOptions opts) =>
        {
            BenchmarkDotNet.Running.BenchmarkRunner.Run<TestSingleBenchmark>();
            return 0;
        },
        errs => 1);
