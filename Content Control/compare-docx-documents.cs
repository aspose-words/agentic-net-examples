using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents from disk.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count > 0 || docEdited.Revisions.Count > 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Configure comparison options (optional). Here we compare all elements and use the edited document as the target.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,                 // Do not track moved text.
            IgnoreFormatting = false,             // Consider formatting changes.
            IgnoreCaseChanges = false,            // Case‑sensitive comparison.
            IgnoreComments = false,               // Include comment changes.
            IgnoreTables = false,                 // Include table changes.
            IgnoreFields = false,                 // Include field changes.
            IgnoreFootnotes = false,              // Include footnote/endnote changes.
            IgnoreTextboxes = false,              // Include textbox changes.
            IgnoreHeadersAndFooters = false,      // Include header/footer changes.
            Target = ComparisonTargetType.New    // Use the edited document as the base for comparison.
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document that now contains revision marks.
        docOriginal.Save("ComparedResult.docx");
    }
}
