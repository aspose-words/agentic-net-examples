using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the source documents.
        string originalPath = "Original.docx";
        string editedPath   = "Edited.docx";

        // Load the two documents.
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no existing revisions (required by Aspose.Words).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison. The original document will receive revision marks.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the document that now contains highlighted differences.
        string resultPath = "ComparisonResult.docx";
        docOriginal.Save(resultPath);
    }
}
