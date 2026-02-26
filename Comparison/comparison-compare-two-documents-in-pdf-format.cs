using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComparisonExample
{
    static void Main()
    {
        // Paths to the original and edited PDF documents.
        string originalPdfPath = "Original.pdf";
        string editedPdfPath = "Edited.pdf";

        // Load the two PDF files as Aspose.Words Document objects.
        Document originalDoc = new Document(originalPdfPath);
        Document editedDoc = new Document(editedPdfPath);

        // Ensure both documents have no existing revisions before comparison.
        if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
        }

        // Save the comparison result (original document with revisions) as a PDF.
        // PdfSaveOptions can be used to control PDF output; default options are sufficient here.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        originalDoc.Save("ComparisonResult.pdf", pdfOptions);
    }
}
