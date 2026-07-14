using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // If DisplayText is null or empty, use the Url as the link text.
        builder.Writeln("<<if [string.IsNullOrEmpty(item.DisplayText)]>>");
        builder.Writeln("<<link [item.Url] [item.Url]>>");
        builder.Writeln("<</if>>");

        // If DisplayText has a value, use it as the link text.
        builder.Writeln("<<if [!string.IsNullOrEmpty(item.DisplayText)]>>");
        builder.Writeln("<<link [item.Url] [item.DisplayText]>>");
        builder.Writeln("<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<LinkItem>
            {
                new LinkItem
                {
                    Url = "https://www.example.com",
                    DisplayText = "Example Site"
                },
                new LinkItem
                {
                    Url = "https://www.nodisplay.com",
                    DisplayText = "" // Empty display text; fallback to URL.
                },
                new LinkItem
                {
                    Url = "https://www.nulldisplay.com",
                    DisplayText = null // Null display text; fallback to URL.
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("HyperlinkReport.docx");
    }
}

// Root data model referenced by the template (named "model").
public class ReportModel
{
    public List<LinkItem> Items { get; set; } = new();
}

// Individual item containing a URL and optional display text.
public class LinkItem
{
    public string Url { get; set; } = string.Empty;
    public string? DisplayText { get; set; }
}
