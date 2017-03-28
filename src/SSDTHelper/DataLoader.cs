using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace SSDTHelper
{
  /// <summary>
  /// Import DataTable to Database
  /// </summary>
  public class DataLoader
  {
    /// <summary>
    /// connection string
    /// </summary>
    public string ConnectionString { get; set; }

    public DataLoader()
    {
    }

    /// <summary>
    /// constructor
    /// </summary>
    /// <param name="connectionString">connection string</param>
    public DataLoader(string connectionString)
    {
      this.ConnectionString = connectionString;
    }

    /// <summary>
    /// Import a DataTable to Database
    /// </summary>
    /// <param name="dataTable">the data to import</param>
    /// <param name="withTruncateTable">truncate table before import data. default true.</param>
    public void Load(DataTable dataTable, bool withTruncateTable = true)
    {
      if (string.IsNullOrEmpty(ConnectionString))
      {
        throw new InvalidOperationException("ConnectionString must not be null or empty.");
      }
      if (dataTable == null)
      {
        throw new ArgumentNullException("dataTable must not be null.", "dataTable");
      }
      if (withTruncateTable)
      {
        TruncateTable(dataTable.TableName);
      }

      using (var bulkCopy = new SqlBulkCopy(this.ConnectionString, SqlBulkCopyOptions.KeepIdentity))
      {
        foreach (var c in dataTable.Columns)
        {
          var name = ((DataColumn)c).ColumnName;
          bulkCopy.ColumnMappings.Add(new SqlBulkCopyColumnMapping(name, name));
        }
        bulkCopy.DestinationTableName = dataTable.TableName;
        bulkCopy.WriteToServer(dataTable);
      }
    }

    /// <summary>
    /// Import DataTables to Database
    /// </summary>
    /// <param name="dataTables">the data to import</param>
    /// <param name="withTruncateTable">truncate table before import data. default true.</param>
    public void Load(IEnumerable<DataTable> dataTables, bool withTruncateTable = true)
    {
      foreach (var dataTable in dataTables)
      {
        Load(dataTable, withTruncateTable);
      }
    }

    private void TruncateTable(string tableName)
    {
      var escapedTableName = tableName.Replace("'", "''");
      using (var cn = new SqlConnection(this.ConnectionString))
      {
        cn.Open();
        using (var cmd = cn.CreateCommand())
        {
          var sql = "BEGIN " + 
                   $"  DECLARE @sql NVARCHAR(MAX) = 'TRUNCATE TABLE ' + QUOTENAME('{escapedTableName}') " +
                    "  EXECUTE(@sql) " +
                    "END";
          cmd.CommandText = sql;
          cmd.CommandType = CommandType.Text;
          cmd.ExecuteNonQuery();
        }
      }
    }

    public DataTable GetSchema(string tableName)
    {
      using (var cn = new SqlConnection(this.ConnectionString))
      {
        cn.Open();
        using (var cmd = cn.CreateCommand())
        {
          var sql = $"SELECT TOP 0 * FROM {tableName}";

          cmd.CommandText = sql;
          cmd.CommandType = CommandType.Text;

          var dt = new DataTable();
          dt.TableName = tableName;
          dt.Load(cmd.ExecuteReader());

          return dt;
        }
      }
    }
  }
}
