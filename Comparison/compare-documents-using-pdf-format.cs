using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparisonPdf
{
    static void Main()
    {
        // Paths to the source PDF documents.
        string originalPdfPath = @"C:\Docs\Original.pdf";
        string editedPdfPath   = @"C:\Docs\Edited.pdf";

        // Load the two PDF documents.
        Document docOriginal = new Document(originalPdfPath);
        Document docEdited   = new Document(editedPdfPath);

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Prepare PDF save options (optional: set compliance, etc.).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: set PDF/A-1b compliance.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the comparison result as a PDF file.
        string resultPdfPath = @"C:\Docs\ComparisonResult.pdf";
        docOriginal.Save(resultPdfPath, pdfOptions);
    }
}
