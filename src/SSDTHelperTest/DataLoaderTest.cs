using Dapper;
using NUnit.Framework;
using System.Linq;

namespace SSDTHelperTest
{

  [TestFixture]
  public class DataLoaderTest
  {
 
    [OneTimeSetUp]
    public void Init()
    {

    }

    [Test]
    public void Load_SingleTable()
    {
      var loader = new SSDTHelper.DataLoader();
      loader.ConnectionString = Config.ConnectionString;

      const string sheetName = "People";
      var schema = loader.GetSchema(sheetName);

      var dt = SSDTHelper.ExcelReader.Read(Util.GetLocalFileFullPath("TestData.xlsx"), sheetName, schema);
      loader.Load(dt);

      using (var cn = new System.Data.SqlClient.SqlConnection(Config.ConnectionString))
      {
        cn.Open();
        var sql = @"SELECT * FROM People ORDER BY ID";
        var result = cn.Query(sql).ToList();

        if (result.Count != 3)
        {
          Assert.True(false, $"3行あるはずが{result.Count}行しかない");
          return;
        }


        if (result[0].Id != 1 || result[0].Name != "Hoge" || result[0].Age != 20)
        {
          Assert.True(false, $"1行目のデータが何か違う");
          return;
        }
        if (result[1].Id != 2 || result[1].Name != "Fuga" || result[1].Age != 99)
        {
          Assert.True(false, $"2行目のデータが何か違う");
          return;
        }
        if (result[2].Id != 3 || result[2].Name != "FugaFuga" || result[2].Age != 0)
        {
          Assert.True(false, $"3行目のデータが何か違う");
          return;
        }

        Assert.True(true);
      }
    }

    [Test]
    public void Load_MultiTables()
    {
      var dt = SSDTHelper.ExcelReader.Read(Util.GetLocalFileFullPath("TestData.xlsx"), new string[] { "People", "Salary" });

      var loader = new SSDTHelper.DataLoader();
      loader.ConnectionString = Config.ConnectionString;
      loader.Load(dt);

      using (var cn = new System.Data.SqlClient.SqlConnection(Config.ConnectionString))
      {
        cn.Open();

        var result = cn.Query(@"SELECT * FROM People ORDER BY ID").ToList();

        if (result.Count != 3)
        {
          Assert.True(false, $"3行あるはずが{result.Count}行しかない");
          return;
        }

        if (result[0].Id != 1 || result[0].Name != "Hoge" || result[0].Age != 20)
        {
          Assert.True(false, $"1行目のデータが何か違う");
          return;
        }
        if (result[1].Id != 2 || result[1].Name != "Fuga" || result[1].Age != 99)
        {
          Assert.True(false, $"2行目のデータが何か違う");
          return;
        }
        if (result[2].Id != 3 || result[2].Name != "FugaFuga" || result[2].Age != 0)
        {
          Assert.True(false, $"3行目のデータが何か違う");
          return;
        }

        var result2 = cn.Query(@"SELECT * FROM Salary ORDER BY ID").ToList();

        if (result2.Count != 3)
        {
          Assert.True(false, $"3行あるはずが{result2.Count}行しかない");
          return;
        }


        if (result2[0].Id != 1 || result2[0].Salary != 3000)
        {
          Assert.True(false, $"1行目のデータが何か違う");
          return;
        }
        if (result2[1].Id != 2 || result2[1].Salary != 5000)
        {
          Assert.True(false, $"2行目のデータが何か違う");
          return;
        }
        if (result2[2].Id != 3 || result2[2].Salary != 4000)
        {
          Assert.True(false, $"3行目のデータが何か違う");
          return;
        }

        Assert.True(true);
      }
    }

    [Test]
    public void Load_NullValue()
    {
      var dt = SSDTHelper.ExcelReader.Read(Util.GetLocalFileFullPath("TestData.xlsx"), "NullValue");

      var loader = new SSDTHelper.DataLoader();
      loader.ConnectionString = Config.ConnectionString;
      loader.Load(dt);

      using (var cn = new System.Data.SqlClient.SqlConnection(Config.ConnectionString))
      {
        cn.Open();
        var sql = @"SELECT * FROM NullValue ORDER BY ID";
        var result = cn.Query(sql).ToList();
        
        Assert.IsNull(result[0].IntCol01, $"IntCol01がNULLではない");
        Assert.IsNull(result[0].DecCol01, $"DecCol01がNULLではない");
        Assert.IsNull(result[0].DateCol01, $"DateCol01がNULLではない");
        Assert.IsNull(result[0].DateTimeCol01, $"DateTimeCol01がNULLではない");
        Assert.IsNull(result[0].VarCharCol01, $"VarCharCol01がNULLではない");
      }
    }

    [Test]
    public void Load_SkipBlankRow()
    {
      var dt = SSDTHelper.ExcelReader.Read(Util.GetLocalFileFullPath("TestData.xlsx"), "BlankRow");

      var loader = new SSDTHelper.DataLoader();
      loader.ConnectionString = Config.ConnectionString;
      loader.Load(dt);

      using (var cn = new System.Data.SqlClient.SqlConnection(Config.ConnectionString))
      {
        cn.Open();
        var sql = @"SELECT * FROM BlankRow ORDER BY ID";
        var result = cn.Query(sql).ToList();

        if (result.Count != 3)
        {
          Assert.True(false, $"3行あるはずが{result.Count}行しかない");
          return;
        }


        if (result[0].Id != 1 || result[0].Name != "Hoge")
        {
          Assert.True(false, $"1行目のデータが何か違う");
          return;
        }
        if (result[1].Id != 2 || result[1].Name != "Fuga")
        {
          Assert.True(false, $"2行目のデータが何か違う");
          return;
        }
        if (result[2].Id != 3 || result[2].Name != "FugaFuga")
        {
          Assert.True(false, $"3行目のデータが何か違う");
          return;
        }

        Assert.True(true);
      }
    }
  }
}
