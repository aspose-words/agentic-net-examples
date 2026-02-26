using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

namespace AsposeWordsLinqReporting
{
    // Simple data class used as the data source for the ReportingEngine.
    public class ReportData
    {
        // Made nullable to silence CS8618 warnings (or you could initialise them).
        public string? Title { get; set; }
        public string? HtmlContent { get; set; }
    }

    // Callback that inserts the HTML string at the location of the matched placeholder.
    public class HtmlInsertCallback : IReplacingCallback
    {
        private readonly string _html;

        public HtmlInsertCallback(string html)
        {
            _html = html;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The Document property of MatchNode returns DocumentBase in newer versions.
            // Cast it to Document, which is the type expected by DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
            builder.MoveTo(args.MatchNode);
            // Insert the HTML fragment.
            builder.InsertHtml(_html);
            // Skip the default replacement (the placeholder text itself).
            return ReplaceAction.Skip;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the TXT template. The template contains ReportingEngine tags:
            //   <<[ds.Title]>>
            //   <<[ds.HtmlContent]>>
            string templatePath = "Template.txt";

            // Load the TXT template as a Word document.
            Document doc = new Document(templatePath);

            // Prepare the data source.
            ReportData data = new ReportData
            {
                Title = "Dynamic LINQ Reporting",
                HtmlContent = "<p><b>This paragraph is inserted as HTML.</b> It can contain <i>any</i> HTML markup.</p>"
            };

            // Populate the template with the data source using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" matches the tags in the template.
            engine.BuildReport(doc, data, "ds");

            // After the report is built, replace the HTML placeholder with the actual HTML.
            FindReplaceOptions options = new FindReplaceOptions
            {
                // Use a custom callback to insert HTML instead of performing a plain text replace.
                ReplacingCallback = new HtmlInsertCallback(data.HtmlContent ?? string.Empty)
            };

            // The placeholder that will be replaced.
            string placeholder = "<<[ds.HtmlContent]>>";

            // Perform the replace operation. The callback inserts the HTML and skips the placeholder.
            doc.Range.Replace(placeholder, string.Empty, options);

            // Save the final document.
            doc.Save("Result.docx");
        }
    }
}
