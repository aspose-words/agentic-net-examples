using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of link items.
        public List<LinkItem> Items { get; set; } = new();
    }

    // Individual link item.
    public class LinkItem
    {
        // URL of the hyperlink.
        public string Url { get; set; } = string.Empty;

        // Optional display text; may be empty.
        public string DisplayText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the temporary template and the final report.
            const string templatePath = "template.docx";
            const string outputPath = "report.docx";

            // ---------- Create the template document ----------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // If DisplayText is null or empty, use the URL as the display text.
            builder.Writeln("<<if [string.IsNullOrEmpty(item.DisplayText)]>>");
            builder.Writeln("<<link [item.Url] [item.Url]>>");
            builder.Writeln("<</if>>");

            // If DisplayText has a value, use it as the display text.
            builder.Writeln("<<if [!string.IsNullOrEmpty(item.DisplayText)]>>");
            builder.Writeln("<<link [item.Url] [item.DisplayText]>>");
            builder.Writeln("<</if>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // ---------- Load the template and build the report ----------
            var reportDoc = new Document(templatePath);

            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<LinkItem>
                {
                    new LinkItem { Url = "https://www.example.com", DisplayText = "Example Site" },
                    new LinkItem { Url = "https://www.github.com", DisplayText = "" }, // Empty display text.
                    new LinkItem { Url = "https://www.microsoft.com", DisplayText = null } // Null display text.
                }
            };

            // Create the reporting engine and build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
