using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = "";
    public string? LinkText { get; set; }
}

public class ReportData
{
    public List<ReportModel> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data: one link with custom text, one without.
        var data = new ReportData
        {
            Items = new()
            {
                new ReportModel { Url = "https://example.com/page1", LinkText = "Visit Page 1" },
                new ReportModel { Url = "https://example.com/page2", LinkText = null }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Link Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Link tag with optional display text.
        builder.Writeln("<<link [item.Url] [item.LinkText]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, data, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
