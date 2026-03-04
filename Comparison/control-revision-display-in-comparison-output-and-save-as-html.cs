using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

namespace RevisionComparisonToHtml
{
    class Program
    {
        static void Main()
        {
            // Create the original document and add some content.
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);
            builder.Writeln("This is the original paragraph.");
            builder.Writeln("Another line that will stay unchanged.");

            // Clone the original to create an edited version and modify its content.
            Document docEdited = (Document)docOriginal.Clone(true);
            DocumentBuilder editBuilder = new DocumentBuilder(docEdited);
            // Change the first paragraph.
            Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;
            firstParagraph.Runs[0].Text = "This is the edited paragraph.";
            // Insert a new paragraph.
            editBuilder.MoveToDocumentEnd();
            editBuilder.Writeln("A newly added line.");

            // Ensure both documents have no revisions before comparison.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Compare the documents – revisions will be added to the original document.
            docOriginal.Compare(docEdited, "John Doe", DateTime.Now);

            // Configure how revisions are displayed in the rendered output.
            RevisionOptions revOptions = docOriginal.LayoutOptions.RevisionOptions;
            revOptions.ShowOriginalRevision = true;      // Show original text alongside revisions.
            revOptions.ShowRevisionMarks = true;        // Keep revision markup (e.g., underline/strike).
            revOptions.ShowRevisionBars = false;        // Hide side revision bars.
            revOptions.ShowInBalloons = ShowInBalloons.None; // Render revisions inline, not in balloons.

            // Prepare HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                HtmlVersion = HtmlVersion.Html5,   // Use HTML5 compliance.
                PrettyFormat = true               // Make the output human‑readable.
            };

            // Save the comparison result as an HTML file.
            string outputPath = "RevisionComparison.html";
            docOriginal.Save(outputPath, htmlOptions);

            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
