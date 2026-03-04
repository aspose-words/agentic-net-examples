using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the source DOC files.
        string originalPath = @"C:\Docs\Original.doc";
        string editedPath   = @"C:\Docs\Edited.doc";

        // Load the original and edited documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no tracked changes before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Compare the documents. The revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Optional: accept all revisions so that docOriginal becomes identical to docEdited.
        // docOriginal.Revisions.AcceptAll();

        // Save the result of the comparison.
        string resultPath = @"C:\Docs\Compared.doc";
        docOriginal.Save(resultPath);
    }
}
