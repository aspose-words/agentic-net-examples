using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class RevisionComparisonToHtml
{
    static void Main()
    {
        // Paths to the source documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";

        // Load the original document – this will become the document that holds the revisions.
        Document doc = new Document(originalPath);

        // Load the document to compare against.
        Document docToCompare = new Document(editedPath);

        // Perform the comparison. The result (revisions) will be stored in 'doc'.
        doc.Compare(docToCompare, "Reviewer", DateTime.Now);

        // -----------------------------------------------------------------
        // Configure how revisions are displayed in the rendered output.
        // -----------------------------------------------------------------
        RevisionOptions revOptions = doc.LayoutOptions.RevisionOptions;

        // Show the original text instead of the revised one.
        revOptions.ShowOriginalRevision = true;

        // Keep the revision markup (e.g., strike‑through, underline) visible.
        revOptions.ShowRevisionMarks = true;

        // Hide the side revision bars that appear next to lines with changes.
        revOptions.ShowRevisionBars = false;

        // Render all revisions inline (no balloons).
        revOptions.ShowInBalloons = ShowInBalloons.None;

        // -----------------------------------------------------------------
        // Save the document as HTML with pretty formatting.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Use HTML5 output.
            HtmlVersion = HtmlVersion.Html5,

            // Make the generated HTML human‑readable.
            PrettyFormat = true
        };

        // Save the result.
        string outputPath = @"C:\Docs\ComparisonResult.html";
        doc.Save(outputPath, htmlOptions);
    }
}
