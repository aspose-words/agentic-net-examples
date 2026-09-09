using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    // Name of the item.
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    // Collection of items to be used in the template.
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
                new Item { Name = "Alpha" },
                new Item { Name = "Beta" },
                new Item { Name = "Gamma" }   // This is the last item.
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Show total count.
        builder.Writeln("Total items: <<[model.Items.Count]>>");

        // Use ElementAt with calculated index to display the last item's name.
        // The expression is evaluated by the LINQ Reporting Engine.
        builder.Writeln("Last item: <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        var engine = new ReportingEngine();
        // BuildReport must be called after the template is fully prepared.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
