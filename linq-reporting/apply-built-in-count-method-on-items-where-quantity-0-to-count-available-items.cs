using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a Word document template programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line that uses the built‑in Count method to count items with Quantity > 0.
        builder.Writeln("Available items count: <<[model.Items.Count(i => i.Quantity > 0)]>>");

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Quantity = 5 },
                new Item { Name = "Banana", Quantity = 0 },
                new Item { Name = "Orange", Quantity = 3 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
