using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model for a single hyperlink entry.
    public class LinkInfo
    {
        // URL of the hyperlink.
        public string Url { get; set; } = string.Empty;

        // Optional display text; if empty the URL will be used as the display text.
        public string DisplayText { get; set; } = string.Empty;
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<LinkInfo> Links { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            // Begin a foreach loop over the Links collection.
            builder.Writeln("<<foreach [link in Links]>>");

            // If DisplayText is not empty, use it as the link's display text.
            builder.Writeln("<<if [!string.IsNullOrEmpty(link.DisplayText)]>>");
            builder.Writeln("<<link [link.Url] [link.DisplayText]>>");
            builder.Writeln("<</if>>");

            // If DisplayText is empty, omit the second argument so the URL is used as display text.
            builder.Writeln("<<if [string.IsNullOrEmpty(link.DisplayText)]>>");
            builder.Writeln("<<link [link.Url]>>");
            builder.Writeln("<</if>>");

            // End the foreach loop and add a line break after each link.
            builder.Writeln("<</foreach>>");
            builder.Writeln(); // add an empty paragraph for readability

            // Save the template to disk.
            builder.Document.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Links = new List<LinkInfo>
                {
                    new LinkInfo { Url = "https://www.example.com", DisplayText = "Example Site" },
                    new LinkInfo { Url = "https://www.github.com", DisplayText = "" }, // empty display text
                    new LinkInfo { Url = "https://www.microsoft.com", DisplayText = "Microsoft" }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            // The root object name used in the template is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
