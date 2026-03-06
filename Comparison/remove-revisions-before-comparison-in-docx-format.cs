using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the source documents and the output file.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";
        string resultPath   = "ComparedResult.docx";

        // Load the original and edited documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Remove any existing revisions from both documents before comparison.
        // AcceptAll() clears revisions; you could also use RejectAll() if desired.
        if (docOriginal.Revisions.Count > 0)
            docOriginal.Revisions.AcceptAll();

        if (docEdited.Revisions.Count > 0)
            docEdited.Revisions.AcceptAll();

        // Perform the comparison. The original document will receive the revision marks.
        docOriginal.Compare(docEdited, "Author", DateTime.Now);

        // Save the resulting document (which now contains the comparison revisions).
        docOriginal.Save(resultPath);
    }
}
