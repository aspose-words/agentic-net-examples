using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Model classes used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Collection of items to be iterated over in the template.
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        // URL of the hyperlink.
        public string Url { get; set; } = string.Empty;

        // Optional display text; may be empty.
        public string LinkText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Items.Add(new Item { Url = "https://example.com/first", LinkText = "First Link" });
            model.Items.Add(new Item { Url = "https://example.com/second", LinkText = "" }); // Empty display text.
            model.Items.Add(new Item { Url = "https://example.com/third", LinkText = null! }); // Null treated as empty.

            // Create a new blank document and build the template.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // If LinkText is not empty, use it as the display text.
            builder.Writeln("<<if [!string.IsNullOrEmpty(item.LinkText)]>><<link [item.Url] [item.LinkText]>> <</if>>");

            // If LinkText is empty, fall back to using the URL as the display text.
            builder.Writeln("<<if [string.IsNullOrEmpty(item.LinkText)]>><<link [item.Url] [item.Url]>> <</if>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("Report.docx");
        }
    }
}
