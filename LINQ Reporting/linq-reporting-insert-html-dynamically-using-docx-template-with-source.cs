using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class used by the ReportingEngine.
    public class ReportData
    {
        // HTML fragment that should be inserted into the document.
        public string HtmlFragment { get; set; }

        // When true the HTML will be inserted using the builder's formatting as a base.
        // When false the HTML will keep its own default formatting.
        public bool UseBuilderFormatting { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains a placeholder tag <<[data.HtmlFragment]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the final report will be saved.
            const string outputPath = @"C:\Reports\GeneratedReport.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare the data source.
            var data = new ReportData
            {
                // Example HTML – can be any valid fragment.
                HtmlFragment = "<h2 style=\"color:blue;\">Dynamic Title</h2>" +
                               "<p>This paragraph is <b>bold</b> and <i>italic</i>.</p>",
                // Switch to control whether builder formatting is applied.
                UseBuilderFormatting = true   // Change to false to keep HTML's own formatting.
            };

            // Build the report using the LINQ ReportingEngine.
            var engine = new ReportingEngine();
            // The data source name ("data") matches the placeholder in the template.
            engine.BuildReport(template, data, "data");

            // After the report is built, replace the placeholder with the actual HTML.
            // The placeholder text is the same as the field name used in the template.
            // Find the placeholder run and replace it with the HTML content.
            ReplacePlaceholderWithHtml(template, "[data.HtmlFragment]", data.HtmlFragment, data.UseBuilderFormatting);

            // Save the final document.
            template.Save(outputPath, SaveFormat.Docx);
        }

        /// <summary>
        /// Finds a placeholder text in the document and replaces it with the supplied HTML.
        /// The insertion respects the <paramref name="useBuilderFormatting"/> flag.
        /// </summary>
        private static void ReplacePlaceholderWithHtml(Document doc, string placeholder, string html, bool useBuilderFormatting)
        {
            // Locate the placeholder text.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            foreach (Run run in runs)
            {
                if (run.Text.Contains(placeholder))
                {
                    // Move the builder to the start of the run that contains the placeholder.
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.MoveTo(run);

                    // Remove the placeholder text.
                    run.Text = run.Text.Replace(placeholder, string.Empty);

                    // Insert the HTML with the appropriate options.
                    if (useBuilderFormatting)
                    {
                        // Use builder formatting as the base formatting for the inserted HTML.
                        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting);
                    }
                    else
                    {
                        // Insert HTML with its own default formatting.
                        builder.InsertHtml(html);
                    }

                    // Only replace the first occurrence.
                    break;
                }
            }
        }
    }
}
