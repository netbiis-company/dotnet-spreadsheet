using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Netbiis.Spreadsheet
{
  public class Excel : Spreadsheet
  {


    private static Stylesheet CreateStylesheet()
    {
      Stylesheet ss = new Stylesheet();

      Font font0 = new Font();         // Default font

      Font font1 = new Font();         // Bold font
      Color color1 = new Color();
      color1.Rgb = "FFFFFFFF";
      Bold bold = new Bold();
      font1.Append(bold);
      font1.Append(color1);

      Fonts fonts = new Fonts();      // <APENDING Fonts>
      fonts.Append(font0);
      fonts.Append(font1);

      // <Fills>
      Fill fill0 = new Fill();        // Default fill

      Fill fill1 = new Fill();        // Default fill
      PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
      ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "808080" }  };
      BackgroundColor backgroundColor1 = new BackgroundColor() { Rgb = new HexBinaryValue() { Value = "808080" } };
      patternFill3.Append(foregroundColor1);
      patternFill3.Append(backgroundColor1);
      fill1.Append(patternFill3);

      Fills fills = new Fills();      // <APENDING Fills>
      fills.Append(fill0);
      fills.Append(fill1);

      // <Borders>
      Border border0 = new Border();     // Defualt border

      Borders borders = new Borders();    // <APENDING Borders>
      borders.Append(border0);

      CellFormat cellformat0 = new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }; // Default style : Mandatory | Style ID =0

      CellFormat cellformat1 = new CellFormat() { FontId = 1 , FillId = 1};
      CellFormats cellformats = new CellFormats();
      cellformats.Append(cellformat0);
      cellformats.Append(cellformat1);


      ss.Append(fonts);
      ss.Append(fills);
      ss.Append(borders);
      ss.Append(cellformats);


      return ss;
    }

    /// <summary>
    ///   Generates the file.
    /// </summary>
    /// <returns></returns>
    public override byte[] GenerateFile()
    {
      var values = new List<IEnumerable<object>>();
      values.Add(Header);
      values.AddRange(Body);




      var memory = new MemoryStream();
      using (var spreadsheetDocument = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
      {
        var workbookPart = spreadsheetDocument.AddWorkbookPart();


        // Adding style
        WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylePart.Stylesheet = CreateStylesheet();
        stylePart.Stylesheet.Save();

        var sheetName = "Document";
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        worksheetPart.Worksheet = new Worksheet(sheetData);
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook.AppendChild(new Sheets());

        var sheet = new Sheet
        {
          Id = workbookPart.GetIdOfPart(worksheetPart),
          SheetId = 1,
          Name = sheetName
        };

        for (var i = 0; i < values.Count; i++)
        {
          var value = values[i];
          var row = new Row { RowIndex = (uint)i + 1 };
          foreach (var col in value)

          {
            var cell = ConvertObjectToCell(col);
            cell.StyleIndex = 1;
            row.AppendChild(cell);
          }
         
          sheetData.AppendChild(row);
        }

        workbookPart.Workbook.Sheets.AppendChild(sheet);
        workbookPart.Workbook.Save();
      }


      TextWriter writer = new StreamWriter(memory);
      writer.Flush();
      memory.Position = 0;
      var file = memory.ToArray();
      memory.Dispose();

      return file;
    }

    /// <summary>
    ///   Converts the object to cell.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns></returns>
    private static Cell ConvertObjectToCell(object value)
    {
      var cell = new Cell();

      if (value == null)
      {
        cell.DataType = CellValues.String;
        cell.CellValue = new CellValue(string.Empty);
      }
      else
      {
        var objType = value.GetType();

        if (objType == typeof(decimal))
        {
          cell.DataType = CellValues.Number;
          cell.CellValue = new CellValue(value.ToString());
        }
        else if (objType == typeof(int))
        {
          cell.DataType = CellValues.Number;
          cell.CellValue = new CellValue(value.ToString());
        }
        else
        {
          cell.DataType = CellValues.InlineString;
          var inlineString = new InlineString();
          var text = new Text { Text = value.ToString() };
          inlineString.AppendChild(text);
          cell.AppendChild(inlineString);
        }
      }
      return cell;
    }
  }
}
