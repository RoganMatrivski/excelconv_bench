using System.Data;
using ClosedXML.Excel;

public class ConverterClosedXML : IDisposable
{
    private const int maxRows = 1_048_576;
    // private const int maxRows = 16;
    private readonly DataTable dt;
    private string[] longestStringByColIdx = [];
    private XLWorkbook workbook;

    public ConverterClosedXML(DataTable dt)
    {
        this.dt = dt;
        workbook = new();
    }

    public ConverterClosedXML(int row = 32, int col = 32)
    {
        dt = Generator.GenerateDataTable(row, col);
        workbook = new();
    }

    // I wish i found these before OpenXML
    public ConverterClosedXML Setup()
    {
        var headerRowList = dt.Columns.AsEnumerable().Select(c => c.ColumnName);

        foreach (var (pageIdx, rowChunk) in dt.AsEnumerable().Chunk(maxRows).Indexed())
        {
            var sheet = workbook.AddWorksheet($"Data {pageIdx}");
            sheet.Cell("A1").InsertData(headerRowList);
            sheet.Cell("A2").InsertData(rowChunk);
        }

        longestStringByColIdx = dt.Columns.Cast<DataColumn>()
        .Select(col => dt
                .AsEnumerable()
                .Select(row => row[col]?.ToString() ?? string.Empty)
                .OrderByDescending(str => str?.Length ?? 0)
                .FirstOrDefault())
        .ToArray();

        return this;
    }

    public ConverterClosedXML SetWidth_AutoSizeColumn()
    {
        workbook.Worksheets.ForEach(s => s.Columns().ForEach(c => c.AdjustToContents()));
        return this;
    }

    public ConverterClosedXML SetWidth_SkiaSharp()
    {
        var defaultCharWidth = Helper.GetDefaultCharWidth();
        workbook.Worksheets.ForEach(s => s.Columns().Indexed().ForEach(x =>
        {
            var (i, c) = x;
            var colStr = longestStringByColIdx[i];
            var stringWidth = Helper.GetStringWidth(colStr + new string('0', 6) ?? "", defaultCharWidth);

            c.Width = stringWidth / 256;
        }));

        return this;
    }

    public ConverterClosedXML SetWidth_CharCount()
    {
        var defaultCharWidth = Helper.GetDefaultCharWidth();
        workbook.Worksheets.ForEach(s => s.Columns().Indexed().ForEach(x =>
        {
            var (i, c) = x;
            var colStr = longestStringByColIdx[i];
            var stringWidth = colStr.Length;

            c.Width = stringWidth;
        }));

        return this;
    }

    public void WriteToStream(Stream stream) => workbook.SaveAs(stream);

    public byte[] WriteToByteArray()
    {
        MemoryStream ms = new();
        workbook.SaveAs(ms);

        return ms.ToArray();
    }

    public void Dispose()
    {
        workbook.Dispose();
        GC.SuppressFinalize(this);
    }
}