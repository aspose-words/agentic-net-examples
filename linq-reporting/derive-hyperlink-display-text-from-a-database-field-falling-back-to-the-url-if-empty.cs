using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Links Report");
        builder.Writeln("<<foreach [item in Items]>>");
        // If LinkText is empty, fall back to Url as the display text.
        builder.Writeln("<<link [item.Url] [string.IsNullOrEmpty(item.LinkText) ? item.Url : item.LinkText]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data: some items have LinkText, others rely on Url.
        ReportModel model = new()
        {
            Items = new List<Item>
            {
                new() { Url = "https://example.com/page1", LinkText = "Example Page 1" },
                new() { Url = "https://example.com/page2", LinkText = "" },
                new() { Url = "https://example.com/page3", LinkText = null! },
                new() { Url = "https://example.com/page4", LinkText = "Custom Text 4" }
            }
        };

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
