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
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Score = 92 },
                new Item { Score = 67 },
                new Item { Score = 45 }
            }
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Dynamic Text Color Example");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a paragraph with a textColor tag that uses a conditional expression.
        // The color changes based on the numeric Score value.
        builder.Writeln(
            "<<textColor [item.Score > 80 ? \"Green\" : item.Score > 50 ? \"Orange\" : \"Red\"]>>" +
            "Score: <<[item.Score]>>" +
            "<</textColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for reporting.
        var loadedDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report_Output.docx";
        loadedDoc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item containing a numeric value.
public class Item
{
    public int Score { get; set; }
}
