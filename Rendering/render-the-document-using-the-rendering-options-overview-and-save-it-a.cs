using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Create PDF save options and configure rendering.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Enable high‑quality (slow) rendering for better visual fidelity.
        pdfOptions.UseHighQualityRendering = true;
        // Disable memory optimization (optional, set to true for large docs).
        pdfOptions.MemoryOptimization = false;

        // Save the document as PDF using the specified options.
        doc.Save("output.pdf", pdfOptions);
    }
}
