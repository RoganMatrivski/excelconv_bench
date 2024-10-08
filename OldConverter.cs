using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System.Data;

public class OldConverter
{
    public static byte[] ConvertToExcel(DataTable dt, string UserID)
    {
        int excellColumnMaxLengthValue = 512;

        int maxRows = 1048576;

        IWorkbook workbook = new XSSFWorkbook();

        IDataFormat dataFormatCustom = workbook.CreateDataFormat();
        ICellStyle dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd");

        ISheet sheet = null;

        int counter = 0;
        int page = 0;
        foreach (DataRow dr in dt.Rows)
        {
            if (counter > maxRows - 1)
            {
                counter = 0;
                page++;
            }

            if (counter == 0)
            {
                sheet = workbook.CreateSheet("Data" + (dt.Rows.Count > maxRows - 1 ? " " + (page + 1) : ""));
                IRow headerRow = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    DataColumn dc = dt.Columns[i];
                    ICell cell = headerRow.CreateCell(counter++, CellType.String);

                    cell.SetCellValue(dc.ColumnName);
                }

                counter = 1;
            }

            IRow row = sheet.CreateRow(counter++);

            int columnCounter = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                ICell cell = row.CreateCell(columnCounter++);
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

        // THIS ONE IS PROBLEMATIC
        // Either AutoSizeColumn, GetColumnWidth, or SetColumnWidth causes SIGNIFICANT slow down
        // Tracks down to NPOI.GetCellWidth -> SixLabors.Fonts causes slowdown.
        for (int i = 0; i <= dt.Columns.Count; i++)
        {
            sheet.AutoSizeColumn(i);

            if (sheet.GetColumnWidth(i) > excellColumnMaxLengthValue * 256)
            {
                sheet.SetColumnWidth(i, excellColumnMaxLengthValue * 256);
            }
        }

        MemoryStream ms = new MemoryStream();
        workbook.Write(ms);

        return ms.ToArray();
    }
}
