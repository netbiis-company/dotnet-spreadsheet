using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Netbiis.Spreadsheet;
using Netbiis.SpreadsheetTest.Mock;
using Xunit;

namespace Netbiis.SpreadsheetTest
{
  public class SpreadsheetTest
  {
    private readonly MockObject _mock = new MockObject(25);
    private readonly string _applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

    [Fact]
    public void GenerateCsvFromObjectList()
    {
      Spreadsheet.Spreadsheet spreadsheet = new Csv();
      spreadsheet.SetData(_mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateCsvFromObjectList.csv", file);
    }

    [Fact]
    public void GenerateExcelFromExcelCellListWithStylesheet()
    {
      var spreadsheet = new Excel();
      spreadsheet.Stylesheet = ExcelStylesheet.Default;
      spreadsheet.SetStylesheet(0, 1);
      spreadsheet.SetStylesheet(7, 1);

      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor"};
      spreadsheet.SetDataByExcelCell(header, _mock.ExcelCellItems);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateExcelFromExcelCellListWithStylesheet.xlsx", file);
    }

    [Fact]
    public void GenerateExcelFromObjectList()
    {
      Spreadsheet.Spreadsheet spreadsheet = new Excel();
      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor", "Oi?"};
      spreadsheet.SetData(header, _mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateExcelFromObjectList.xlsx", file);
    }

    [Fact]
    public void GenerateExcelFromObjectListWithStylesheet()
    {
      var spreadsheet = new Excel();
      spreadsheet.Stylesheet = ExcelStylesheet.Default;
      spreadsheet.SetStylesheet(0, 1);
      spreadsheet.SetStylesheet(7, 1);

      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor", "Oi?"};
      spreadsheet.SetData(header, _mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateExcelFromObjectListWithStylesheet.xlsx", file);
    }
  }
}
