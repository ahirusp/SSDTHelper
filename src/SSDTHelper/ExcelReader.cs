using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace SSDTHelper
{
  /// <summary>
  /// Read Excel file to DataTable
  /// </summary>
  public class ExcelReader
  {
    const string COMMENT_COLUMN_HEADER = "**comment**";

    /// <summary>
    /// Reads Excel file and creates DataTable
    /// </summary>
    /// <param name="excellFilePath">The path of the target excel file.</param>
    /// <param name="sheetName">The name of the target sheet.</param>
    /// <returns>DataTable object</returns>
    /// <remarks>
    /// Read a specified sheet of Excel file into DataTable.
    /// The sheet name is set in the TableName property of the DataTable.
    /// </remarks>
    public static DataTable Read(string excellFilePath, string sheetName, DataTable schema = null)
    {
      return Read(excellFilePath, (new string[] { sheetName }), (schema == null) ? null : (new DataTable[] { schema }))[0];
    }

    /// <summary>
    /// Reads Excel file and creates DataTable
    /// </summary>
    /// <param name="excellFilePath">The path of the target excel file.</param>
    /// <param name="sheetNames">The names of the target sheets.</param>
    /// <returns>DataTable objects</returns>
    /// <remarks>
    /// Read a specified sheet of Excel file into DataTable.
    /// The sheet name is set in the TableName property of the DataTable.
    /// </remarks>
    public static IList<DataTable> Read(string path, IEnumerable<string> sheetNames, IEnumerable<DataTable> schemas = null)
    {
      using (var xl = new ExcelPackage())
      {
        var dts = new List<DataTable>();

        using (var st = File.OpenRead(path))
        {
          xl.Load(st);
        }

        foreach (var sheetName in sheetNames)
        {
          using (var ws = xl.Workbook.Worksheets[sheetName])
          {
            if(ws == null)
            {
              throw new ArgumentException($"'{sheetName}’ is not found.");
            }

            var dt = new DataTable() { TableName = sheetName };

            var endColumn = ws.Dimension.End.Column;

            if(schemas != null && schemas.Any((x) => x.TableName == sheetName))
            {
              dt = schemas.First((x) => x.TableName == sheetName).Clone();

              if (ws.Cells[1, endColumn].Text == COMMENT_COLUMN_HEADER)
              {
                endColumn--;
              }
            }
            else
            {
              foreach (var cell in ws.Cells[1, 1, 1, endColumn])
              {
                if (cell.End.Column == ws.Dimension.End.Column && cell.Text == COMMENT_COLUMN_HEADER)
                {
                  endColumn = cell.End.Column - 1;
                  break;
                }
                dt.Columns.Add(cell.Text);
              }
            }

            var startRow = 2;
            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
              var wsRow = ws.Cells[rowNum, 1, rowNum, endColumn];

              if (wsRow.Any(x => !string.IsNullOrEmpty(x.Text)))
              {
                DataRow row = dt.NewRow();
                foreach (var cell in wsRow)
                {
                  if (cell.Text == "NULL")
                  {
                    row[cell.Start.Column - 1] = DBNull.Value;
                  }
                  else if (dt.Columns[cell.Start.Column - 1].DataType == typeof(bool))
                  {
                    row[cell.Start.Column - 1] = cell.Value;
                  }
                  else if (dt.Columns[cell.Start.Column - 1].DataType == typeof(int))
                  {
                    row[cell.Start.Column - 1] = cell.Value;
                  }
                  else if (dt.Columns[cell.Start.Column - 1].DataType == typeof(long))
                  {
                    row[cell.Start.Column - 1] = cell.Value;
                  }
                  else if (dt.Columns[cell.Start.Column - 1].DataType == typeof(double))
                  {
                    row[cell.Start.Column - 1] = cell.Value;
                  }
                  else if (dt.Columns[cell.Start.Column - 1].DataType == typeof(DateTime))
                  {
                    row[cell.Start.Column - 1] = cell.Value;
                  }
                  else
                  {
                    row[cell.Start.Column - 1] = cell.Text;
                  }
                }
                dt.Rows.Add(row);
              }
            }
            dts.Add(dt);
          }

        }
        return dts;
      }
    }
  }
}
