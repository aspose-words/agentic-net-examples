using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfComplianceDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Docs\ResultDocument.pdf";

            // Choose the desired PDF compliance level.
            // Options include PDF/A (e.g., PdfA1b, PdfA2u, PdfA4Ua2) or PDF/UA (e.g., PdfUa1, PdfUa2).
            PdfCompliance compliance = PdfCompliance.PdfA1b; // Change as needed.

            // Load the Word document (creation/loading rule).
            Document doc = new Document(inputPath);

            // Configure PDF save options (creation rule).
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // Apply the selected compliance level.
                Compliance = compliance
            };

            // Save the document as PDF with the specified compliance (saving rule).
            doc.Save(outputPath, saveOptions);
        }
    }
}
