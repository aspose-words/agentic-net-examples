using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new LinkInfo
        {
            Url = "https://example.com",
            DisplayText = "Example Site"
        };

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting link tags.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Link that specifies both the target URL and the display text.
        builder.Writeln("Link with explicit display text:");
        builder.Writeln("<<link [model.Url] [model.DisplayText]>>");

        // Link that specifies only the target URL.
        // The engine will use the URL itself as the display text.
        builder.Writeln("Link with default display text (URL will be shown):");
        builder.Writeln("<<link [model.Url]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model object. The root name is "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class LinkInfo
{
    // Target URL for the hyperlink (required).
    public string Url { get; set; } = string.Empty;

    // Optional display text for the hyperlink.
    // If null or empty, the link tag without a second expression will be used.
    public string? DisplayText { get; set; }
}
