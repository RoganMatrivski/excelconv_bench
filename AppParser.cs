using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System.Linq;

public static class AppParser
{
    public static int RunConvertAndReturnExitCode(Args.ConvertOptions opts)
    {
        DataTable dataTable = new DataTable();
        using (StreamReader sr = new StreamReader(opts.Path))
        {
            string[]? headers = sr.ReadLine()?.Split(',');
            foreach (string header in headers)
            {
                dataTable.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[]? rows = sr.ReadLine()?.Split(',');
                DataRow dr = dataTable.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dataTable.Rows.Add(dr);
            }
        }

        if (dataTable.Rows.Count <= 0)
        {
            throw new Exception("No data found in the file");
        }


        using (var workbook = SpreadsheetDocument.Create(opts.OutputPath, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            // Add header row
            var headerRow = new Row();
            foreach (DataColumn column in dataTable.Columns)
            {
                var cell = new Cell { DataType = CellValues.String, CellValue = new CellValue(column.ColumnName) };
                headerRow.AppendChild(cell);
            }
            sheetData.AppendChild(headerRow);

            // Add data rows
            foreach (DataRow dataRow in dataTable.Rows)
            {
                var newRow = new Row();
                foreach (var item in dataRow.ItemArray)
                {
                    var cell = new Cell { DataType = CellValues.String, CellValue = new CellValue(item.ToString()) };
                    newRow.AppendChild(cell);
                }
                sheetData.AppendChild(newRow);
            }


            workbookPart.Workbook.Save();
        }

        Console.WriteLine($"Excel file created successfully at: {opts.OutputPath}");

        return 0;
    }

    public static byte[] ConvertTable(DataTable dataTable)
    {
        var outputStream = new MemoryStream();

        if (dataTable.Rows.Count <= 0)
        {
            throw new Exception("No data found in the file");
        }


        var defaultCharWidth = Helper.GetDefaultCharWidth();

        var longestStringByColumns = dataTable.Columns.Cast<DataColumn>()
        .Select(col => dataTable
                .AsEnumerable()
                .Select(row => row[col]?.ToString() ?? string.Empty)
                .OrderByDescending(str => str?.Length ?? 0)
                .FirstOrDefault())
        .ToList();

        using (var workbook = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();

            // Set column width
            var columns = new Columns();
            for (uint i = 1; i <= dataTable.Columns.Count; i++) // Assuming you want to set width for each column
            {
                // Measure string padded with some char for some padding space
                // It won't be visible in the final result, i promise
                var stringWidth = Helper.GetStringWidth(longestStringByColumns[Convert.ToInt32(i - 1)] + new string('0', 6) ?? "", defaultCharWidth);
                // Console.WriteLine(Math.Min(256, stringWidth / 256));

                var column = new Column
                {
                    Min = i,
                    Max = i,
                    Width = Math.Min(256, stringWidth / 256), // Width in characters
                    CustomWidth = true
                };
                columns.Append(column);
            }

            worksheetPart.Worksheet = new Worksheet(columns, sheetData);

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            // Add header row
            var headerRow = new Row();
            foreach (DataColumn column in dataTable.Columns)
            {
                var cell = new Cell { DataType = CellValues.String, CellValue = new CellValue(column.ColumnName) };
                headerRow.AppendChild(cell);
            }
            sheetData.AppendChild(headerRow);

            // Add data rows
            foreach (DataRow dataRow in dataTable.Rows)
            {
                var newRow = new Row();
                foreach (var item in dataRow.ItemArray)
                {
                    var cell = new Cell { DataType = CellValues.String, CellValue = new CellValue(item.ToString()) };
                    newRow.AppendChild(cell);
                }
                sheetData.AppendChild(newRow);
            }

            // for (int i = 0; i < dataTable.Columns.Count; i++)
            // foreach (var (i, col) in sheet.Elements<Column>().Select((c, i) => (i, c)))
            // {
            //     var txtSize = Helper.GetStringWidth(longestStringByColumns[i] ?? "", defaultCharWidth);

            //     col.Width = Math.Min(txtSize, 100 * 256);
            //     col.CustomWidth = true;
            // }

            foreach (var c in sheet.Elements<Column>())
            {
                Console.WriteLine($"Col width: {c.Width}");
            }

            workbookPart.Workbook.Save();
        }

        return outputStream.ToArray();
    }
}
