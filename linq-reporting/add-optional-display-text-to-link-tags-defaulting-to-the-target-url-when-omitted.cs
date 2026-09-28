using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the LINQ Reporting template.
        builder.Writeln("Link examples:");
        builder.Writeln("<<foreach [item in Items]>>");
        // Link with explicit display text.
        builder.Writeln("- <<link [item.Url] [item.LinkText]>>");
        // Link without display text – should default to the URL.
        builder.Writeln("- <<link [item.Url]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, can be omitted if building directly).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Load the template back (demonstrates load step).
        Document templateDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item
                {
                    Url = "https://example.com/page1",
                    LinkText = "Example Page 1"
                },
                new Item
                {
                    Url = "https://example.com/page2",
                    LinkText = null // No display text; engine will use the URL.
                }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        templateDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the foreach loop.
public class Item
{
    public string Url { get; set; } = "";
    public string? LinkText { get; set; }
}
