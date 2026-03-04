using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and edited DOC files.
        Document docOriginal = new Document(@"C:\Docs\Original.doc");
        Document docEdited   = new Document(@"C:\Docs\Edited.doc");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");
        }

        // Compare the documents. The original document will receive Revision nodes for each change.
        docOriginal.Compare(docEdited, "AuthorInitials", DateTime.Now);

        // Optional: accept all revisions to transform the original into the edited version.
        // docOriginal.Revisions.AcceptAll();

        // Save the result (original document now contains tracked changes) as a new DOC file.
        docOriginal.Save(@"C:\Docs\ComparedResult.doc");
    }
}
