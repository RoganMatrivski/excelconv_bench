using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// You have no idea how much time spent 
// to even understand OpenXML Excel structure
// Even with a guide, i barely grasp it
// Guide: https://jason-ge.medium.com/create-excel-using-openxml-in-net-6-3b601ddf48f7
public class ConverterOpenXML : IDisposable
{
    private const int maxRows = 1_048_576;
    // private const int maxRows = 16;
    private readonly DataTable dt;
    private string[] longestStringByColIdx = [];
    private SpreadsheetDocument workbook;
    private Columns workbookColumns;
    private Action columnSetupFn;
    private Stream output;

    public ConverterOpenXML(Stream output, int row = 32, int col = 32)
    {
        this.output = output;
        dt = Generator.GenerateDataTable(row, col);
        workbook = SpreadsheetDocument.Create(output, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
    }

    public ConverterOpenXML(Stream output, DataTable dt)
    {
        this.dt = dt;
        workbook = SpreadsheetDocument.Create(output, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
    }

    public ConverterOpenXML SetSetupColumns_SkiaSharp()
    {
        columnSetupFn = () =>
        {
            var defaultCharWidth = Helper.GetDefaultCharWidth();

            var columns = new Columns();
            for (uint i = 1; i <= dt.Columns.Count; i++) // Assuming you want to set width for each column
            {
                // Measure string padded with some char for some padding space
                // It won't be visible in the final result, i promise
                var stringWidth = Helper.GetStringWidth(longestStringByColIdx[Convert.ToInt32(i - 1)] + new string('0', 6) ?? "", defaultCharWidth);

                var column = new Column
                {
                    Min = i,
                    Max = i,
                    Width = Math.Min(256, stringWidth / 256), // Width in characters
                    CustomWidth = true
                };
                columns.Append(column);
            }

            workbookColumns = columns;
        };

        return this;
    }

    public ConverterOpenXML SetSetupColumns_CharCount()
    {
        columnSetupFn = () =>
        {
            var columns = new Columns();
            for (uint i = 1; i <= dt.Columns.Count; i++) // Assuming you want to set width for each column
            {
                // Measure string padded with some char for some padding space
                // It won't be visible in the final result, i promise
                var stringWidth = longestStringByColIdx[Convert.ToInt32(i - 1)].Length;

                var column = new Column
                {
                    Min = i,
                    Max = i,
                    Width = Math.Min(256, stringWidth), // Width in characters
                    CustomWidth = true
                };
                columns.Append(column);
            }

            workbookColumns = columns;
        };

        return this;
    }

    public ConverterOpenXML ExportTable()
    {
        var dtCol = dt.Columns;
        var colEnum = dtCol.AsEnumerable();
        var colEnumIdx = colEnum.Indexed();

        longestStringByColIdx = colEnum
        .Select(col => dt
                .AsEnumerable()
                .Select(row => row[col]?.ToString() ?? string.Empty)
                .OrderByDescending(str => str?.Length ?? 0)
                .FirstOrDefault())
        .ToArray();

        columnSetupFn();

        var workbookPart = workbook.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var headerRow = new Row(colEnum.Select(col => new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(col.ColumnName)
        }));

        foreach (var (pageIdx, rowChunk) in dt.AsEnumerable().Chunk(maxRows).Indexed())
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = Convert.ToUInt32(pageIdx + 1),
                Name = $"Data {pageIdx + 1}"
            };
            sheets.Append(sheet);

            var sheetData = new SheetData(
                rowChunk.Select(dr =>
                    new Row(colEnum.Select(dc => DataColumnToCell(dr, dc)))
                ).Prepend(headerRow.CloneNode(true))
            );

            worksheetPart.Worksheet = new(workbookColumns.CloneNode(true), sheetData);
        }

        workbookPart.Workbook.Save();

        return this;
    }

    private Cell DataColumnToCell(DataRow dr, DataColumn dc)
    {
        switch (Type.GetTypeCode(dc.DataType))
        {
            case TypeCode.Int16:
            case TypeCode.Int32:
            case TypeCode.Int64:
            case TypeCode.Single:
            case TypeCode.Double:
            case TypeCode.Decimal:
                return new Cell { DataType = CellValues.Number, CellValue = new CellValue(Convert.ToDouble(dr[dc.ColumnName])) };
            case TypeCode.Boolean:
                return new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(Convert.ToBoolean(dr[dc.ColumnName])) };
            case TypeCode.DateTime:
                return new Cell { DataType = CellValues.Date, CellValue = new CellValue(Convert.ToDateTime(dr[dc.ColumnName])) };
            case TypeCode.String:
                string data = Convert.ToString(dr[dc.ColumnName]);
                return new Cell { DataType = CellValues.String, CellValue = new CellValue(data.Length > 32767 ? data.Substring(0, 32767) : data) };
            default:
                return new Cell { DataType = CellValues.String, CellValue = new CellValue(Convert.ToString(dr[dc.ColumnName])) };
        }
    }

    public void Dispose()
    {
        workbook.Dispose();
    }
}