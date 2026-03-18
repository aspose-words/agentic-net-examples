using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = "Original.docx";
        string editedPath = "Edited.docx";
        string outputPath = "Compared.docx";

        // Load the original document.
        Document docOriginal = new Document(originalPath);

        // Load the edited document.
        Document docEdited = new Document(editedPath);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Perform the comparison.
        // Author name and comparison date are specified here.
        string author = "John Doe";
        DateTime comparisonDate = DateTime.Now;
        docOriginal.Compare(docEdited, author, comparisonDate);

        // Save the document that now contains revision differences.
        docOriginal.Save(outputPath);
    }
}
