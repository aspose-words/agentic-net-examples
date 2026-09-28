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
            Items = new List<Item>
            {
                new() { Text = "First centered paragraph.", Alignment = "center" },
                new() { Text = "Second right‑aligned paragraph.", Alignment = "right" },
                new() { Text = "Third left‑aligned paragraph.", Alignment = "left" }
            }
        };

        // Create the template document.
        Document template = new();
        DocumentBuilder builder = new(template);

        // Write a foreach block that outputs each item with dynamic alignment using an HTML tag.
        builder.Writeln("<<foreach [item in Items]>>");
        // The HTML expression builds a <p> element with the desired text‑align style.
        builder.Writeln("<<html [\"<p style='text-align:\" + item.Alignment + \";'>\" + item.Text + \"</p>\"]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new(templatePath);
        ReportingEngine engine = new()
        {
            Options = ReportBuildOptions.None
        };

        // Build the report.
        engine.BuildReport(report, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Text { get; set; } = string.Empty;
    public string Alignment { get; set; } = "left";
}
