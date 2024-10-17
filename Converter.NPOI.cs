using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;

public class ConverterNPOI
{
    private const int maxRows = 1_048_576;
    // private const int maxRows = 16;

    private readonly DataTable dt;
    private string[] longestStringByColIdx = [];
    // private ISheet? sheet;
    private IWorkbook workbook;

    public ConverterNPOI(int row = 32, int col = 32)
    {
        dt = Generator.GenerateDataTable(row, col);
        workbook = new XSSFWorkbook();
    }

    public ConverterNPOI(DataTable dt)
    {
        this.dt = dt;
        workbook = new XSSFWorkbook();
    }

    public ConverterNPOI SetupSheet()
    {
        var dataFormatCustom = workbook.CreateDataFormat();
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd");

        var columnEnum = dt.Columns.AsEnumerable();

        // Chunk the rows to max rows
        foreach (var (pageIdx, rowChunk) in dt.AsEnumerable().Chunk(maxRows).Indexed())
        {
            var chunkSheet = workbook.CreateSheet($"Data {pageIdx + 1}".TrimEnd());

            foreach (var (rowIdx, dr) in rowChunk.Indexed())
            {
                var headerRow = chunkSheet.CreateRow(0);

                foreach (var (colIdx, dc) in columnEnum.Indexed())
                {
                    headerRow
                        .CreateCell(colIdx, CellType.String)
                        .SetCellValue(dc.ColumnName);
                }

                var row = chunkSheet.CreateRow(rowIdx);

                foreach (var (colIdx, dc) in columnEnum.Indexed())
                {
                    var cell = row.CreateCell(colIdx);

                    if (dc.DataType == typeof(DateTime))
                    {
                        cell.CellStyle = dateStyle;
                    }

                    if (dr[dc.ColumnName] == DBNull.Value) continue;

                    switch (Type.GetTypeCode(dc.DataType))
                    {
                        case TypeCode.Int16:
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Single:
                        case TypeCode.Double:
                            cell.SetCellValue(Convert.ToDouble(dr[dc.ColumnName]).ToString());
                            break;
                        case TypeCode.Boolean:
                            cell.SetCellValue(Convert.ToBoolean(dr[dc.ColumnName]));
                            break;
                        case TypeCode.DateTime:
                            cell.SetCellValue(Convert.ToDateTime(dr[dc.ColumnName]));
                            break;
                        case TypeCode.String:
                            string data = Convert.ToString(dr[dc.ColumnName]);
                            cell.SetCellValue(data.Length > 32767 ? data.Substring(0, 32767) : data);
                            break;
                        default:
                            cell.SetCellValue(Convert.ToString(dr[dc.ColumnName]));
                            break;
                    }
                }
            }
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

    public ConverterNPOI SetWidth_AutoSizeColumn()
    {
        for (var i = 0; i < dt.Columns.Count; i++)
        {
            foreach (var sheet in workbook)
            {
                sheet.AutoSizeColumn(i);
            }
        }

        return this;
    }

    public ConverterNPOI SetWidth_SkiaSharp()
    {
        var defaultCharWidth = Helper.GetDefaultCharWidth();
        for (var i = 0; i < dt.Columns.Count; i++)
        {
            var colStr = longestStringByColIdx[i];
            var stringWidth = Helper.GetStringWidth(colStr + new string('0', 6) ?? "", defaultCharWidth);

            foreach (var sheet in workbook)
            {
                sheet.SetColumnWidth(i, Math.Min(255 * 256, stringWidth));
            }
        }

        return this;
    }

    public ConverterNPOI SetWidth_CharCount()
    {
        for (var i = 0; i < dt.Columns.Count; i++)
        {
            var colStr = longestStringByColIdx[i];
            var stringWidth = colStr.Length;

            foreach (var sheet in workbook)
            {
                sheet.SetColumnWidth(i, Math.Min(255, stringWidth) * 255);
            }
        }

        return this;
    }

    public void WriteToStream(Stream stream) => workbook.Write(stream);

    public byte[] WriteToByteArray()
    {
        MemoryStream ms = new();
        workbook.Write(ms);

        return ms.ToArray();
    }
}