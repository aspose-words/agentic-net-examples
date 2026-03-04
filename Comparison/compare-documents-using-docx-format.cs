using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the original and edited documents from disk.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Ensure both documents have no existing revisions; otherwise Compare will throw.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Optional: iterate through the revisions and output their types.
        foreach (Revision rev in docOriginal.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, node type: {rev.ParentNode.NodeType}");
        }

        // Accept all revisions so that docOriginal becomes identical to docEdited.
        docOriginal.Revisions.AcceptAll();

        // Save the resulting document.
        docOriginal.Save("ComparedResult.docx");
    }
}
