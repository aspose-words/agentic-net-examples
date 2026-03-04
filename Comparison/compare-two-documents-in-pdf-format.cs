using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Paths to the PDF files to compare and the output file.
        string originalPdfPath = "original.pdf";
        string editedPdfPath = "edited.pdf";
        string outputPdfPath = "comparison_result.pdf";

        // Load the two PDF documents into Aspose.Words Document objects.
        Document docOriginal = new Document(originalPdfPath);
        Document docEdited = new Document(editedPdfPath);

        // Ensure both documents have no revisions before performing the comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions describing the differences are added to docOriginal.
            docOriginal.Compare(docEdited, "Author", DateTime.Now);
        }

        // Save the original document (now containing revisions) as a PDF file.
        docOriginal.Save(outputPdfPath);
    }
}
