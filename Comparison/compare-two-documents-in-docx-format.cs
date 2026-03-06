using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original document.
        Document original = new Document("Original.docx");

        // Load the document to compare against.
        Document edited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions (required by Aspose.Words).
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(edited, "AuthorInitials", DateTime.Now);

        // Optional: accept all revisions to transform the original into the edited version.
        // original.Revisions.AcceptAll();

        // Save the result containing the tracked changes.
        original.Save("ComparisonResult.docx");
    }
}
