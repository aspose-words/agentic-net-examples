using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source PDF documents.
            const string originalPdfPath = @"C:\Docs\Original.pdf";
            const string editedPdfPath   = @"C:\Docs\Edited.pdf";

            // Load the original and edited documents.
            Document docOriginal = new Document(originalPdfPath);
            Document docEdited   = new Document(editedPdfPath);

            // Ensure both documents have no existing revisions; otherwise Compare will throw.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

            // Compare the documents. The revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

            // Save the comparison result as a PDF file.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set compliance to PDF/A‑1b for archival quality.
                Compliance = PdfCompliance.PdfA1b
            };

            const string resultPdfPath = @"C:\Docs\ComparisonResult.pdf";
            docOriginal.Save(resultPdfPath, pdfOptions);
        }
    }
}
