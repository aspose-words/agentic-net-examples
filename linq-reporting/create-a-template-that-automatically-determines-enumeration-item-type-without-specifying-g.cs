using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a foreach tag that iterates over the Items collection.
        // The engine will infer the item type (Item) automatically.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[item.Index]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items =
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model exposed to the template.
public class ReportModel
{
    // Collection of items; the engine will determine the item type at runtime.
    public List<Item> Items { get; set; } = new();
}

// Simple item class referenced inside the foreach loop.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
