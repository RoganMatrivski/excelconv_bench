using System.Data;

public static class Generator
{
    public static int RunGenerateAndReturnExitCode(Args.GenerateOptions opts)
    {
        var faker = new Bogus.Faker();
        var csvData = new System.Text.StringBuilder();

        // Generate header row
        var headers = new List<string>();
        for (int i = 0; i < opts.Columns; i++)
        {
            headers.Add($"Column{i + 1}");
        }
        csvData.AppendLine(string.Join(",", headers));

        // Generate data rows and write to file
        using (var writer = new StreamWriter(opts.Path))
        {
            // Write header row
            writer.WriteLine(string.Join(",", headers));

            // Generate and write data rows
            int chunkSize = 32;
            int numChunks = (int)Math.Ceiling((double)opts.Rows / chunkSize);

            Parallel.For(0, numChunks, chunkIndex =>
            {
                var localFaker = new Bogus.Faker(); // Create a local Faker instance for thread safety
                var chunkData = new List<string>();

                int startRow = chunkIndex * chunkSize;
                int endRow = Math.Min(startRow + chunkSize, opts.Rows);

                for (int row = startRow; row < endRow; row++)
                {
                    var rowData = new List<string>();
                    for (int col = 0; col < opts.Columns; col++)
                    {
                        switch (col % 4)
                        {
                            case 0:
                                rowData.Add(localFaker.Random.Number(1, 1000).ToString());
                                break;
                            case 1:
                                rowData.Add(localFaker.Random.Bool().ToString());
                                break;
                            case 2:
                                rowData.Add(localFaker.Date.Past().ToString("yyyy-MM-dd"));
                                break;
                            case 3:
                                rowData.Add($"{localFaker.Lorem.Sentence(3).Replace(",", "")}");
                                break;
                        }
                    }
                    chunkData.Add(string.Join(",", rowData));
                }

                lock (writer)
                {
                    foreach (var row in chunkData)
                    {
                        writer.WriteLine(row);
                    }
                }
            });
        }
        Console.WriteLine($"Generated CSV file with {opts.Rows} rows and {opts.Columns} columns.");
        return 0;
    }

    public static DataTable GenerateDataTable(int rows, int columns)
    {
        var faker = new Bogus.Faker();
        var dataTable = new DataTable();

        // Add columns
        for (int i = 0; i < columns; i++)
        {
            switch (i % 4)
            {
                case 0:
                    dataTable.Columns.Add($"Column{i + 1}", typeof(decimal));
                    break;
                case 1:
                    dataTable.Columns.Add($"Column{i + 1}", typeof(bool));
                    break;
                case 2:
                    dataTable.Columns.Add($"Column{i + 1}", typeof(DateTime));
                    break;
                case 3:
                    dataTable.Columns.Add($"Column{i + 1}", typeof(string));
                    break;
            }
        }

        // Generate rows
        for (int i = 0; i < rows; i++)
        {
            var row = dataTable.NewRow();
            for (int j = 0; j < columns; j++)
            {
                switch (j % 4)
                {
                    case 0:
                        row[j] = faker.Random.Decimal(0, 1000);
                        break;
                    case 1:
                        row[j] = faker.Random.Bool();
                        break;
                    case 2:
                        row[j] = faker.Date.Past().ToString("dd MMM yyyy");
                        break;
                    case 3:
                        row[j] = faker.Lorem.Sentence();
                        break;
                }
            }
            dataTable.Rows.Add(row);
        }

        return dataTable; ;
    }
}
