using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class DocumentComparison
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document original = new Document("Original.docx");
        Document edited   = new Document("Edited.docx");

        // Set up comparison options (default values are sufficient for a full comparison).
        CompareOptions compareOptions = new CompareOptions
        {
            // Use the original document as the base for comparison.
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. The revisions are added to the 'original' document.
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

        // Configure how revisions will be rendered in the output.
        RevisionOptions revOptions = original.LayoutOptions.RevisionOptions;
        revOptions.InsertedTextColor   = RevisionColor.Green;          // Inserted text in green.
        revOptions.InsertedTextEffect  = RevisionTextEffect.Underline; // Underline inserted text.
        revOptions.DeletedTextColor    = RevisionColor.Red;            // Deleted text in red.
        revOptions.DeletedTextEffect   = RevisionTextEffect.StrikeThrough; // Strike‑through deletions.
        revOptions.ShowRevisionMarks   = true;                         // Show revision marks.
        revOptions.ShowOriginalRevision = false;                       // Show revised text, not original.

        // Save the result as HTML with revisions highlighted.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            HtmlVersion = HtmlVersion.Html5,
            PrettyFormat = true
        };

        original.Save("ComparisonResult.html", htmlOptions);
    }
}
