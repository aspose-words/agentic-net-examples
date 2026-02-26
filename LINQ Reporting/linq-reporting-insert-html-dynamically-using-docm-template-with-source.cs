using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data source class used by the ReportingEngine.
    public class ReportData
    {
        // The HTML fragment that will be inserted into the document.
        public string HtmlContent { get; set; }

        // Determines whether the builder's formatting should be applied to the inserted HTML.
        public bool UseBuilderFormatting { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCM template that contains a bookmark named "HtmlPlaceholder".
            // The template can also contain other LINQ Reporting placeholders that will be
            // populated automatically by the ReportingEngine.
            const string templatePath = @"C:\Templates\ReportTemplate.docm";

            // Load the DOCM template.
            Document doc = new Document(templatePath);

            // Prepare the data source.
            var data = new ReportData
            {
                HtmlContent = "<h2 style=\"color:blue;\">Dynamic Title</h2>" +
                              "<p>This paragraph is inserted <b>as HTML</b> using LINQ Reporting.</p>",
                UseBuilderFormatting = true // Switch: true => apply builder formatting, false => default HTML formatting.
            };

            // Build the report using the LINQ Reporting engine.
            // The data source name "src" is used inside the template to reference members,
            // e.g. <<[src.HtmlContent]>> or <<[src.UseBuilderFormatting]>>.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "src");

            // After the report is built, insert the HTML fragment at the bookmark.
            // The bookmark must exist in the template; otherwise MoveToBookmark will throw.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToBookmark("HtmlPlaceholder");

            // Switch based on the sourceStyles flag (UseBuilderFormatting).
            // When true, the builder's current formatting (e.g., paragraph alignment, font) is used
            // as the base formatting for the imported HTML.
            // When false, the HTML's own styling is preserved.
            if (data.UseBuilderFormatting)
            {
                // Example: set a base paragraph alignment that will be applied to HTML elements
                // lacking explicit alignment.
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.InsertHtml(data.HtmlContent, true); // useBuilderFormatting = true
            }
            else
            {
                // No builder formatting is applied; HTML renders as it would in a browser.
                builder.InsertHtml(data.HtmlContent, false); // useBuilderFormatting = false
            }

            // Save the resulting document.
            const string outputPath = @"C:\Output\GeneratedReport.docx";
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
