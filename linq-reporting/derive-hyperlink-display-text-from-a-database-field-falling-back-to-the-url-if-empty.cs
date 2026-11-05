using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        System.IO.Directory.CreateDirectory("output");

        // Step 1: Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple heading.
        builder.Writeln("Hyperlink Report");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a hyperlink where the display text falls back to the URL if LinkText is empty or null.
        builder.Writeln("<<link [item.Url] [(!string.IsNullOrEmpty(item.LinkText)) ? item.LinkText : item.Url]>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "output/template.docx";
        templateDoc.Save(templatePath);

        // Step 2: Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Step 3: Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<LinkItem>
            {
                new LinkItem { Url = "https://example.com/1", LinkText = "Example One" },
                new LinkItem { Url = "https://example.com/2", LinkText = "" },
                new LinkItem { Url = "https://example.com/3", LinkText = null }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        string outputPath = "output/HyperlinkReport.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<LinkItem> Items { get; set; } = new();
}

// Individual item containing a URL and optional display text.
public class LinkItem
{
    public string Url { get; set; } = string.Empty;
    public string? LinkText { get; set; }
}
