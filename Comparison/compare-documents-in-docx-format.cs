using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the original, edited, and result documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";
        string resultPath   = @"C:\Docs\ComparisonResult.docx";

        // Load the two documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no tracked revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Optional: accept all revisions so that docOriginal becomes identical to docEdited.
        // docOriginal.Revisions.AcceptAll();

        // Save the document that now contains the revision markup.
        docOriginal.Save(resultPath);
    }
}
