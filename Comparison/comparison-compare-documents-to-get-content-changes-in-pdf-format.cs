using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparisonToPdf
{
    static void Main()
    {
        // Paths to the original and edited documents.
        string originalPath = @"C:\Docs\Original.docx";
        string editedPath   = @"C:\Docs\Edited.docx";

        // Load the documents using the Document constructor (lifecycle rule).
        Document docOriginal = new Document(originalPath);
        Document docEdited   = new Document(editedPath);

        // Ensure both documents have no existing revisions (required before comparison).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Perform the comparison. The original document will receive revision marks.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the result as PDF. The PDF will display the tracked changes.
        string outputPdfPath = @"C:\Docs\ComparisonResult.pdf";
        docOriginal.Save(outputPdfPath); // Extension determines PDF format.
    }
}
