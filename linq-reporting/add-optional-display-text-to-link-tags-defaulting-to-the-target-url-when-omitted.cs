using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // ---------- Create the LINQ Reporting template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Links collection.
        builder.Writeln("<<foreach [link in Links]>>");

        // If LinkText is null or empty, use a link tag without a display text (defaults to URL).
        builder.Writeln("<<if [string.IsNullOrEmpty(link.LinkText)]>>");
        builder.Writeln("<<link [link.Url]>>");
        builder.Writeln("<</if>>");

        // If LinkText has a value, include it as the display text.
        builder.Writeln("<<if [!string.IsNullOrEmpty(link.LinkText)]>>");
        builder.Writeln("<<link [link.Url] [link.LinkText]>>");
        builder.Writeln("<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load the template and build the report ----------
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Links = new List<LinkInfo>
            {
                new LinkInfo
                {
                    Url = "https://example.com",
                    LinkText = "Example Site"
                },
                new LinkInfo
                {
                    Url = "https://aspose.com",
                    LinkText = null // No display text; will default to the URL.
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<LinkInfo> Links { get; set; } = new();
}

// Represents a single hyperlink entry.
public class LinkInfo
{
    public string Url { get; set; } = string.Empty;
    public string? LinkText { get; set; }
}
