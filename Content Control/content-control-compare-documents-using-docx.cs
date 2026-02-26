using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareContentControl
{
    static void Main()
    {
        // Load the original and the edited documents that contain content controls.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options.
            CompareOptions options = new CompareOptions
            {
                // Track changes at the word level.
                Granularity = Granularity.WordLevel,
                // Use the edited document as the target (equivalent to Word's "Show changes in New document").
                Target = ComparisonTargetType.New,
                // Do not ignore any element types; content controls are compared as regular nodes.
                IgnoreFormatting = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                CompareMoves = false
            };

            // Perform the comparison. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, options);
        }

        // Save the document that now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");

        // Optionally accept all revisions to turn the original into the edited version.
        docOriginal.Revisions.AcceptAll();
        docOriginal.Save("ComparisonResult_Accepted.docx");
    }
}
