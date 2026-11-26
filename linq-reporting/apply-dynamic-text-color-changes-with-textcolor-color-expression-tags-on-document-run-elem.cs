using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    // Text to display.
    public string Text { get; set; } = string.Empty;
    // Color expression accepted by the textColor tag (e.g., "Red", "#FF0000").
    public string ColorName { get; set; } = string.Empty;
}

public class ReportModel
{
    // Collection of items to be rendered.
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
                new() { Text = "Success", ColorName = "Green" },
                new() { Text = "Warning", ColorName = "Orange" },
                new() { Text = "Error",   ColorName = "Red" },
                new() { Text = "Info",    ColorName = "#0000FF" } // Hex color.
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Dynamic Text Color Report");
        builder.Writeln("--------------------------");
        builder.Writeln("<<foreach [item in Items]>>");
        // Apply dynamic text color using the textColor tag.
        builder.Writeln("<<textColor [item.ColorName]>><<[item.Text]>><</textColor>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
