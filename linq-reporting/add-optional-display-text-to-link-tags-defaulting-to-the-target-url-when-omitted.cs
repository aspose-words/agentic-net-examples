using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "LinkTemplate.docx";
        const string reportPath = "LinkReport.docx";

        // ---------- Create the LINQ Reporting template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple heading.
        builder.Writeln("Link Report");
        builder.Writeln();

        // Begin a foreach loop over the collection "Links".
        builder.Writeln("<<foreach [link in Links]>>");

        // Insert a link tag. The second expression (display text) is optional.
        // If LinkText is null or empty, the engine will display the URL itself.
        builder.Writeln("<<link [link.Url] [link.LinkText]>>");
        builder.Writeln();

        // End the foreach block.
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
                    Url = "https://www.example.com",
                    LinkText = "Example Site"
                },
                new LinkInfo
                {
                    Url = "https://www.aspose.com",
                    LinkText = null // No display text; URL will be used as the link text.
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
        reportDoc.Save(reportPath);
    }
}

// Wrapper class that matches the root object name used in the template ("model").
public class ReportModel
{
    public List<LinkInfo> Links { get; set; } = new();
}

// Data class representing a single link.
public class LinkInfo
{
    public string Url { get; set; } = string.Empty;
    public string? LinkText { get; set; }
}
