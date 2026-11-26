using System;
using System.Collections.Generic;
using System.IO;
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
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 5 },
                new Item { Name = "Banana", Quantity = 0 },
                new Item { Name = "Orange", Quantity = 3 },
                new Item { Name = "Grapes", Quantity = -1 }
            }
        };

        // Create a template document with a LINQ Reporting tag that counts items where Quantity > 0.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Available items count: <<[model.Items.Count(i => i.Quantity > 0)]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "Report.docx");
        template.Save(outputPath);
    }
}
