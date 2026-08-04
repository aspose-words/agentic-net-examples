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
            Items = new List<ParagraphItem>
            {
                new ParagraphItem { Text = "First centered paragraph.", Alignment = "center" },
                new ParagraphItem { Text = "Second right‑aligned paragraph.", Alignment = "right" },
                new ParagraphItem { Text = "Third left‑aligned paragraph.", Alignment = "left" }
            }
        };

        // Create the template document programmatically.
        string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Items]>>");
        // Use an HTML expression to control alignment dynamically.
        builder.Writeln("<<[p.Html] -html>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");
        if (!success)
        {
            Console.WriteLine("Report generation failed.");
            return;
        }

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}

// Root data model.
public class ReportModel
{
    public List<ParagraphItem> Items { get; set; } = new();
}

// Item model with dynamic alignment.
public class ParagraphItem
{
    public string Text { get; set; } = string.Empty;
    public string Alignment { get; set; } = "left";

    // Returns an HTML snippet that sets the paragraph alignment.
    public string Html => $"<p style=\"text-align:{Alignment}; margin:0;\">{System.Net.WebUtility.HtmlEncode(Text)}</p>";
}
