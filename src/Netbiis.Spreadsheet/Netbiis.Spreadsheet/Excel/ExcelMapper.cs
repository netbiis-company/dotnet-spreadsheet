// ReSharper disable CheckNamespace

using System;
using System.Collections.Generic;

namespace Netbiis.Spreadsheet
{
  public class ExcelMapper<T>
  {
    public ExcelMapper()
    {
      Items = new List<Item>();
    }

    public List<Item> Items { get; set; }

    public void Add(string referenceLetter, string name, Type type)
    {
      Items.Add(new Item
      {
        ReferenceLetter = referenceLetter,
        Name = name,
        Type = type
      });
    }

    public class Item
    {
      public string Name { get; set; }
      public string ReferenceLetter { get; set; }
      public Type Type { get; set; }
    }
  }
}
