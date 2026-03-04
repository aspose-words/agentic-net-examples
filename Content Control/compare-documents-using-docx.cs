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
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            // If revisions exist, reject them before comparison.
            docOriginal.Revisions.RejectAll();
            docEdited.Revisions.RejectAll();
        }

        // Perform the comparison. The original document will receive Revision objects
        // describing the differences found in the edited document.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Optional: accept all revisions so the original document becomes identical to the edited one.
        // Comment out this line if you want to keep the revisions for review.
        docOriginal.Revisions.AcceptAll();

        // Save the result to a new file.
        docOriginal.Save("ComparedResult.docx");
    }
}
