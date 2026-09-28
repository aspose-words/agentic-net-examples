using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);

        // Step 1: Build the LINQ Reporting template programmatically.
        string templatePath = System.IO.Path.Combine(outputDir, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Add a heading.
        builder.Writeln("Dynamic Text Color Report");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply dynamic text color using the textColor tag.
        // The color expression comes from item.Color, and the displayed text from item.Text.
        builder.Writeln("<<textColor [item.Color]>><<[item.Text]>><</textColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Step 2: Load the template for report generation.
        var doc = new Document(templatePath);

        // Step 3: Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Text = "Important", Color = "Red" },
                new Item { Text = "Info", Color = "Green" },
                new Item { Text = "Note", Color = "#0000FF" } // Blue using hex code.
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated report.
        string reportPath = System.IO.Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Public data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Text { get; set; } = string.Empty;
    public string Color { get; set; } = string.Empty;
}
