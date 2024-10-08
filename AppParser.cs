using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

        using (var workbook = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
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
		
        return outputStream.ToArray();
    }
}
