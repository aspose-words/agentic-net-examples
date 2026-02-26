using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfComplianceDemo
{
    class Program
    {
        static void Main()
        {
            // Input document path (replace with your actual file).
            string inputPath = Path.Combine(@"C:\Docs", "Sample.docx");

            // Output folder (replace with your actual folder).
            string outputFolder = @"C:\Output";

            // Load the source document.
            Document doc = new Document(inputPath);

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set compliance to PDF/A-4 and PDF/UA-2.
            // This combines long‑term preservation (PDF/A) with accessibility (PDF/UA).
            pdfOptions.Compliance = PdfCompliance.PdfA4Ua2;

            // Optional: ensure document structure is exported (required for PDF/A‑4 and PDF/UA‑2).
            // The property is ignored for standards that already require it, but setting it
            // makes the intent explicit.
            pdfOptions.ExportDocumentStructure = true;

            // Define the output file name.
            string outputPath = Path.Combine(outputFolder, "Sample.PdfA4Ua2.pdf");

            // Save the document as PDF with the specified compliance.
            doc.Save(outputPath, pdfOptions);

            // Demonstrate saving the same document as PDF/A‑1b (visual preservation only).
            PdfSaveOptions pdfA1bOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };
            string pdfA1bPath = Path.Combine(outputFolder, "Sample.PdfA1b.pdf");
            doc.Save(pdfA1bPath, pdfA1bOptions);

            // Demonstrate saving the same document as PDF/UA‑1 (accessibility only).
            PdfSaveOptions pdfUa1Options = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfUa1
            };
            string pdfUa1Path = Path.Combine(outputFolder, "Sample.PdfUa1.pdf");
            doc.Save(pdfUa1Path, pdfUa1Options);
        }
    }
}
