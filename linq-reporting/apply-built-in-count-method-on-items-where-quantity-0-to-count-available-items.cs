using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = string.Empty;
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
        // Prepare sample data.
        var model = new ReportModel();
        model.Items.Add(new Item { Name = "Apple", Quantity = 5 });
        model.Items.Add(new Item { Name = "Banana", Quantity = 0 });
        model.Items.Add(new Item { Name = "Cherry", Quantity = 12 });
        model.Items.Add(new Item { Name = "Date", Quantity = -3 });

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Insert a tag that counts items with Quantity > 0.
        builder.Writeln("Available items: <<[Items.Count(i => i.Quantity > 0)]>>");
        doc.Save(templatePath);

        // Load the template and build the report.
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        loadedDoc.Save("Report.docx");
    }
}
