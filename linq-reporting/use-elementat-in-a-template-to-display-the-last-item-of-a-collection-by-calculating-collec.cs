using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
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
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Items list:");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("Last item (using ElementAt): <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
