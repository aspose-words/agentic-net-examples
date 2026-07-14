using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple data model.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice" },
                new Item { Name = "Bob" },
                new Item { Name = "Charlie" }
            }
        };

        // Build the template document programmatically.
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        // This tag references a non‑existent member; it will be treated as null.
        builder.Writeln("Missing: <<[item.NonExisting]>>");
        builder.Writeln("<</foreach>>");

        Document doc = builder.Document;

        // Configure the reporting engine to treat missing members as null.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "" // Empty string results in a null literal.
        };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "ReportOutput.docx");
        doc.Save(outputPath);
    }
}

// Root data model containing a collection of items.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class with only a Name property.
public class Item
{
    public string Name { get; set; } = string.Empty;
}
