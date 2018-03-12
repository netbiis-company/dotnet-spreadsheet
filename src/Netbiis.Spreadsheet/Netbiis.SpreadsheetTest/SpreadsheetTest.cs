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

    [Fact]
    public void SpreadsheetReadFile()
    {
      var excelMapper = new ExcelMapper<HotelProvider>();
      excelMapper.Add("A", "HotelKey", typeof(string));
      excelMapper.Add("B", "HotelName", typeof(string));
      excelMapper.Add("C", null, typeof(string));
      excelMapper.Add("D", "LocationType", typeof(string));
      excelMapper.Add("E", "Address_1", typeof(string));
      excelMapper.Add("F", "Address_2", typeof(string));
      excelMapper.Add("G", "Address_3", typeof(string));
      excelMapper.Add("H", "Address_4", typeof(string));
      excelMapper.Add("I", "Phone", typeof(string));
      excelMapper.Add("J", "Fax", typeof(string));
      excelMapper.Add("K", "Email", typeof(string));
      excelMapper.Add("L", "Website", typeof(string));
      excelMapper.Add("M", "StarRating", typeof(int));
      excelMapper.Add("N", "Category", typeof(string));
      excelMapper.Add("O", "Latitude", typeof(double));
      excelMapper.Add("P", "Longitude", typeof(double));
      excelMapper.Add("Q", "CityCode", typeof(string));
      excelMapper.Add("R", "CityName", typeof(string));
      excelMapper.Add("S", "CountryCode", typeof(string));
      excelMapper.Add("T", null, typeof(string));
      excelMapper.Add("U", "CountryName", typeof(string));
      
      var fileName = @"E:\Source\Netbiis-Git\dotnet-spreadsheet\src\Netbiis.Spreadsheet\Netbiis.SpreadsheetTest\Files\HotelProvider.xlsx";
      var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
      var spreadsheet = new Excel();
      var hotelProvider = spreadsheet.ReadFile(fileStream, excelMapper, true);
    }
  }
}
