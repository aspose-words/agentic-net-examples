using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents from files.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count > 0 || docEdited.Revisions.Count > 0)
        {
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");
        }

        // Optional: configure comparison options (e.g., ignore formatting, case changes, etc.).
        CompareOptions compareOptions = new CompareOptions
        {
            // Set desired flags here. For this example we compare everything.
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.Current   // Use the original document as the base.
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "AuthorInitials", DateTime.Now, compareOptions);

        // Save the document that now contains revision marks.
        docOriginal.Save("ComparedResult.docx");
    }
}
