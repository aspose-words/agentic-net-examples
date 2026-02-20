using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;
using Aspose.Words.Comparing;

class RevisionComparisonToHtml
{
    static void Main()
    {
        // Load the two documents that will be compared.
        Document originalDoc = new Document("Original.docx");
        Document editedDoc = new Document("Edited.docx");

        // Configure comparison options (default values are sufficient for this example).
        CompareOptions compareOptions = new CompareOptions
        {
            // Use the edited document as the target so that revisions are created in the original.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to originalDoc.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now, compareOptions);

        // Access the RevisionOptions that control how revisions are rendered.
        RevisionOptions revOptions = originalDoc.LayoutOptions.RevisionOptions;

        // Show the original text alongside the revised text.
        revOptions.ShowOriginalRevision = true;

        // Keep the revision markup (e.g., underlines, strike‑through) visible.
        revOptions.ShowRevisionMarks = true;

        // Render format revisions in balloons on the right side of the page.
        revOptions.ShowInBalloons = ShowInBalloons.Format;

        // Optionally change colors/effects for inserted and deleted text.
        revOptions.InsertedTextColor = RevisionColor.Green;
        revOptions.InsertedTextEffect = RevisionTextEffect.Italic;
        revOptions.DeletedTextColor = RevisionColor.Red;
        revOptions.DeletedTextEffect = RevisionTextEffect.Bold;

        // Save the resulting document (with revisions) as HTML.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Use HTML5 for modern browsers.
            HtmlVersion = HtmlVersion.Html5,

            // Export revision balloons as part of the HTML.
            ExportXhtmlTransitional = false,

            // Optional: pretty‑format the output for readability.
            PrettyFormat = true
        };

        originalDoc.Save("ComparisonResult.html", htmlOptions);
    }
}
