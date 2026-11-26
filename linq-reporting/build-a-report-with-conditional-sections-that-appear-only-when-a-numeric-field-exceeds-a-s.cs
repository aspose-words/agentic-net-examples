using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Alpha", Value = 750 },
                new Item { Name = "Beta", Value = 1250 },
                new Item { Name = "Gamma", Value = 500 },
                new Item { Name = "Delta", Value = 2000 }
            }
        };

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("=== Sample Report ===");
        builder.Writeln("");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- Item: <<[item.Name]>>");
        builder.Writeln("  Value: <<[item.Value]>>");
        builder.Writeln("  <<if [item.Value > 1000]>>High value!<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the report.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
