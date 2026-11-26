using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Url = "https://example.com", LinkText = "Example Site" },
                new Item { Url = "https://github.com", LinkText = "GitHub" },
                new Item { Url = "https://docs.microsoft.com", LinkText = "Microsoft Docs" }
            }
        };

        // Create the template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin the foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build the table inside the loop.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Link");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<link [item.Url] [item.LinkText]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}
