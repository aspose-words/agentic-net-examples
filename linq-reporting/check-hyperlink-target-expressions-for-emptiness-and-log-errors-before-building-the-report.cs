using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportItem
{
    public string Url { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with Hyperlinks:");
        builder.Writeln("<<foreach [item in Items]>>");
        // Link tag: first expression is the URI, second is the display text.
        builder.Writeln("<<link [item.Url] [item.Text]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare sample data.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Items = new List<ReportItem>
            {
                new ReportItem { Url = "https://example.com", Text = "Example Site" },
                new ReportItem { Url = "", Text = "Missing URL" },
                new ReportItem { Url = "https://openai.com", Text = "OpenAI" }
            }
        };

        // Validate hyperlink targets before building the report.
        foreach (var item in model.Items)
        {
            if (string.IsNullOrWhiteSpace(item.Url))
            {
                Console.WriteLine($"Error: Hyperlink target is empty for item with text '{item.Text}'.");
            }
        }

        // Remove items with empty URLs to avoid runtime exception from the link tag.
        model.Items = model.Items
            .Where(i => !string.IsNullOrWhiteSpace(i.Url))
            .ToList();

        // -----------------------------------------------------------------
        // Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        Console.WriteLine(success
            ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
            : "Report generation failed.");
    }
}
