using System;
using Aspose.Words;

class VerifyDocumentAfterRejectingRevisions
{
    static void Main()
    {
        // Load the document that contains tracked changes (revisions).
        Document docWithRevisions = new Document("OriginalWithRevisions.docx");

        // Reject all revisions so the document reverts to its original state.
        docWithRevisions.Revisions.RejectAll();

        // Load the baseline document that represents the expected original content.
        Document baselineDoc = new Document("Baseline.docx");

        // Compare the textual content of both documents.
        bool contentsMatch = docWithRevisions.GetText() == baselineDoc.GetText();

        // Output the verification result.
        Console.WriteLine(contentsMatch
            ? "The document matches the baseline after rejecting revisions."
            : "The document does NOT match the baseline after rejecting revisions.");
    }
}
