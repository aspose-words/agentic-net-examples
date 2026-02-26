using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Paths to the source documents and the output file.
        string originalPath = @"Original.docx";
        string editedPath   = @"Edited.docx";
        string resultPath   = @"Result.docx";

        // Load the original and edited documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no pending revisions before comparison.
        // Accepting all revisions clears the revision collection.
        if (docOriginal.Revisions.Count > 0)
            docOriginal.Revisions.AcceptAll();

        if (docEdited.Revisions.Count > 0)
            docEdited.Revisions.AcceptAll();

        // Perform the comparison. The author name and timestamp are required.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the document that now contains the revision markup.
        docOriginal.Save(resultPath);
    }
}
