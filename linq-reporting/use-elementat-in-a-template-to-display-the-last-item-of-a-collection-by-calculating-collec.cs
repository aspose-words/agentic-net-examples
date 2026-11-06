using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
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
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", Value = 10 },
                new Item { Name = "Beta", Value = 20 },
                new Item { Name = "Gamma", Value = 30 }
            }
        };

        // Create the template document.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("LINQ Reporting Example");
        builder.Writeln("Last item in collection: <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        templateDoc.Save(outputPath);
    }
}
