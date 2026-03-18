using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace PdfAConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            string inputPdfPath = @"C:\Input\source.pdf";

            // Path where the searchable PDF/A‑1a file will be saved.
            string outputPdfPath = @"C:\Output\searchable_a1a.pdf";

            // Load the PDF document. PdfLoadOptions can be used to customize loading,
            // but for this scenario the default options are sufficient.
            Document pdfDocument = new Document(inputPdfPath, new PdfLoadOptions());

            // Configure PDF save options to produce a PDF/A‑1a compliant file.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // PDF/A‑1a requires the document structure (tags) to be present,
                // which makes the content searchable.
                Compliance = PdfCompliance.PdfA1a
            };

            // Save the document as PDF/A‑1a. Aspose.Words will perform OCR on any
            // raster images in the PDF during the save operation, embedding the
            // recognized text so that the resulting file is searchable.
            pdfDocument.Save(outputPdfPath, saveOptions);
        }
    }
}
