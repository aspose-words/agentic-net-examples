using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 5 },
                new Item { Name = "Banana", Quantity = 0 },
                new Item { Name = "Orange", Quantity = 3 }
            }
        };

        // Create a template document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // LINQ Reporting tag that counts items with Quantity > 0.
        builder.Writeln("Available items count: <<[model.Items.Count(i => i.Quantity > 0)]>>");

        // Build the report using the template and the data model.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class used in the collection.
public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
