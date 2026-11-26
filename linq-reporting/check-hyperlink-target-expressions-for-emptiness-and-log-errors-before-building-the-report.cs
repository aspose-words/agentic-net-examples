using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportItem
{
    public string Url { get; set; } = "";
    public string Text { get; set; } = "";
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

        // Sample data.
        var items = new List<ReportItem>
        {
            new() { Url = "https://example.com/page1", Text = "Example Page 1" },
            new() { Url = "", Text = "Missing URL" }, // Triggers validation error.
            new() { Url = "https://example.com/page3", Text = "Example Page 3" }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Hyperlink Report");
        builder.Writeln("<<foreach [item in Items]>>");

        // Output a proper link only when the URL is not empty.
        builder.Writeln("<<if [!string.IsNullOrWhiteSpace(item.Url)]>><<link [item.Url] [item.Text]>><</if>>");
        // For empty URLs, output the display text with a notice.
        builder.Writeln("<<if [string.IsNullOrWhiteSpace(item.Url)]>>[item.Text] (missing URL)<</if>>");

        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // -----------------------------------------------------------------
        // Validate hyperlink targets before building the report.
        // -----------------------------------------------------------------
        foreach (var item in items)
        {
            if (string.IsNullOrWhiteSpace(item.Url))
            {
                Console.WriteLine($"Error: Hyperlink target is empty for display text \"{item.Text}\".");
            }
        }

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        var model = new ReportModel { Items = items };
        bool success = engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        template.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to \"{outputPath}\".");
    }
}
