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

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Url = "https://example.com/page1", Text = "Example Page 1" },
                new Item { Url = "https://example.com/page2", Text = "Example Page 2" },
                new Item { Url = "https://example.com/page3", Text = "Example Page 3" }
            }
        };

        // Create a new document and build the LINQ Reporting template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Dynamic Links Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Insert a link tag with URL and display text.
        builder.Writeln("<<link [item.Url] [item.Text]>>");
        // Add a line break after each link.
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputPath = "DynamicLinksReport.docx";
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model representing a hyperlink.
public class Item
{
    public string Url { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}
