using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<LinkItem>
                {
                    new LinkItem { Url = "https://example.com/page1", DisplayText = "Example Page 1" },
                    new LinkItem { Url = "https://example.com/page2", DisplayText = "" }, // Empty display text.
                    new LinkItem { Url = "https://example.com/page3", DisplayText = null }, // Null display text.
                    new LinkItem { Url = "https://example.com/page4", DisplayText = "Page Four" }
                }
            };

            // Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Generated Links:");
            builder.Writeln("<<foreach [item in Items]>>");
            // If DisplayText is not empty, use it as the link text.
            builder.Writeln(@"<<if [item.DisplayText != """"]>><<link [item.Url] [item.DisplayText]>>><</if>>");
            // If DisplayText is empty (or null), fall back to using the URL as the link text.
            builder.Writeln(@"<<if [item.DisplayText == """"]>><<link [item.Url] [item.Url]>>><</if>>");
            builder.Writeln("<</foreach>>");

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            doc.Save("HyperlinkReport.docx");
        }
    }

    // Root wrapper class for the report.
    public class ReportModel
    {
        public List<LinkItem> Items { get; set; } = new();
    }

    // Data model representing each hyperlink entry.
    public class LinkItem
    {
        public string Url { get; set; } = "";
        public string DisplayText { get; set; } = "";
    }
}
