using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinkItem
{
    public string Url { get; set; } = "";
    public string DisplayText { get; set; } = "";
}

public class ReportModel
{
    public List<LinkItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<LinkItem>
            {
                new LinkItem { Url = "https://example.com/first", DisplayText = "First Site" },
                new LinkItem { Url = "https://example.com/second", DisplayText = "" }, // Empty display text.
                new LinkItem { Url = "https://example.com/third", DisplayText = "Third Site" }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // If DisplayText is empty, use the URL as the link text.
        builder.Writeln("<<if [string.IsNullOrEmpty(item.DisplayText)]>>");
        builder.Writeln("<<link [item.Url] [item.Url]>>");
        builder.Writeln("<</if>>");

        // If DisplayText is not empty, use it as the link text.
        builder.Writeln("<<if [!string.IsNullOrEmpty(item.DisplayText)]>>");
        builder.Writeln("<<link [item.Url] [item.DisplayText]>>");
        builder.Writeln("<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("HyperlinkReport.docx");
    }
}
