using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
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
      File.WriteAllBytes(_applicationPath + @"\file.csv", file);
    }

    [Fact]
    public void GenerateExcelFromObjectList()
    {
      Spreadsheet.Spreadsheet spreadsheet = new Excel();
      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor", "Oi?"};
      spreadsheet.SetData(header, _mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\file.xlsx", file);
    }

    [Fact]
    public void GenerateExcelFromObjectListWithStylesheet()
    {
      Spreadsheet.Excel spreadsheet = new Excel();

      Stylesheet CreateStylesheet()
      {
        var stylesheet = new Stylesheet();

        var font0 = new Font(); // font Default

        var font1 = new Font(); // font 1
        var color1 = new Color {Rgb = "FFFFFFFF"};
        var bold1 = new Bold();
        font1.Append(bold1);
        font1.Append(color1);

        var fonts = new Fonts(); // <APENDING Fonts>
        fonts.Append(font0);
        fonts.Append(font1);

        // <Fills>
        var fill0 = new Fill(); // fill Default
        var fill1 = new Fill(); // fill Default 1

        //var fill2 = new Fill(); // fill 1
        //var patternFill2 = new PatternFill { PatternType = PatternValues.Solid };
        //var foregroundColor2 = new ForegroundColor { Rgb = new HexBinaryValue { Value = "808080" } };
        //patternFill2.Append(foregroundColor2);
        //fill1.Append(patternFill2);


        // <APENDING Fills>
        var fills = new Fills();
        fills.Append(fill0);
        fills.Append(fill1);

        // <Borders>
        var border0 = new Border(); // border Defualt
        var border1 = new Border(); // border Defualt

        // <APENDING Borders>
        var borders = new Borders();
        borders.Append(border0);
        borders.Append(border1);

        var cellformat0 = new CellFormat {FontId = 0, FillId = 0, BorderId = 0}; // Default style : Mandatory | Style ID =0
        var cellformat1 = new CellFormat {FontId = 1, FillId = 0, BorderId = 1 };
        var cellformats = new CellFormats();
        cellformats.Append(cellformat0);
        cellformats.Append(cellformat1);

        stylesheet.Append(fonts);
        //stylesheet.Append(fills);
        stylesheet.Append(borders);
        stylesheet.Append(cellformats);

        return stylesheet;
      }

      spreadsheet.SetStylesheet(0, 1);
      spreadsheet.SetStylesheet(7, 1);
      spreadsheet.Stylesheet = CreateStylesheet();

      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor", "Oi?"};
      spreadsheet.SetData(header, _mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\file.xlsx", file);
    }
  }
}
