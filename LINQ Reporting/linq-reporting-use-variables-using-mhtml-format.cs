using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a template that uses variables and a foreach loop.
        builder.Writeln("Report generated on <<[ReportDate]:date>>");
        builder.Writeln("Total items: <<[ItemsCount]>>");
        builder.Writeln("First item name: <<[FirstItem.Name]>>");
        builder.Writeln("All item names:");
        // The foreach syntax iterates over the Items collection.
        builder.Writeln("<<foreach [in Items]>><<[Name]>> <<end>>");

        // Prepare a data source using LINQ.
        List<Item> items = new List<Item>
        {
            new Item { Name = "Apple",  Price = 1.20 },
            new Item { Name = "Banana", Price = 0.80 },
            new Item { Name = "Cherry", Price = 2.50 }
        };

        var data = new
        {
            ReportDate = DateTime.Now,
            ItemsCount = items.Count,
            FirstItem = items.First(),
            Items = items
        };

        // Build the report by populating the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        // Empty string for dataSourceName because we reference members directly.
        engine.BuildReport(doc, data, "");

        // Save the populated document as MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save("Report.mhtml", saveOptions);
    }

    // Simple POCO class used as an item in the collection.
    public class Item
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }
}
