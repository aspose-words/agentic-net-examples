using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the source documents and the output document.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";
        string resultPath   = "Compared.docx";

        // Load the original and edited documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no tracked revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision-free before comparison.");

        // Perform the comparison. The revisions (differences) will be added to docOriginal.
        // Author initials and timestamp are required parameters.
        docOriginal.Compare(docEdited, "AU", DateTime.Now);

        // Save the document that now contains the highlighted differences.
        docOriginal.Save(resultPath);
    }
}
