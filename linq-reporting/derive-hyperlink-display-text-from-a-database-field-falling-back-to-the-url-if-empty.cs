using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Links = new List<LinkItem>
            {
                new LinkItem { Url = "https://example.com/page1", LinkText = "Example Page 1" },
                new LinkItem { Url = "https://example.com/page2", LinkText = "" }, // Empty display text, should fallback to URL.
                new LinkItem { Url = "https://example.com/page3", LinkText = null } // Null display text, should fallback to URL.
            }
        };

        // Create the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a foreach loop over the Links collection.
        builder.Writeln("<<foreach [link in Links]>>");
        // Insert a link tag where the display text falls back to the URL if LinkText is null or empty.
        builder.Writeln("<<link [link.Url] [string.IsNullOrEmpty(link.LinkText) ? link.Url : link.LinkText]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<LinkItem> Links { get; set; } = new();
}

// Individual link item.
public class LinkItem
{
    public string Url { get; set; } = "";
    public string? LinkText { get; set; }
}
