using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;
using Aspose.Words.Comparing;

class RevisionComparisonToHtml
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";

        // Load the documents.
        Document originalDoc = new Document(originalPath);
        Document editedDoc   = new Document(editedPath);

        // Ensure both documents have no revisions before comparison.
        if (originalDoc.Revisions.Count != 0 || editedDoc.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Set up comparison options (optional – here we compare all elements).
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. The revisions will be added to the original document.
        originalDoc.Compare(editedDoc, "Reviewer", DateTime.Now, compareOptions);

        // Configure how revisions are displayed in the rendered output.
        RevisionOptions revOptions = originalDoc.LayoutOptions.RevisionOptions;
        revOptions.ShowOriginalRevision = true;   // Show original text instead of revised text.
        revOptions.ShowRevisionMarks = true;     // Keep revision markup visible.
        revOptions.ShowRevisionBars = false;     // Hide side revision bars.
        revOptions.ShowInBalloons = ShowInBalloons.None; // Render revisions inline.

        // Prepare HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            HtmlVersion = HtmlVersion.Html5, // Use HTML5 standard.
            PrettyFormat = true            // Make the output human‑readable.
            // Headers and footers are exported by default; no explicit property needed.
        };

        // Save the comparison result as an HTML file.
        string outputPath = @"C:\Docs\ComparisonResult.html";
        originalDoc.Save(outputPath, htmlOptions);
    }
}
