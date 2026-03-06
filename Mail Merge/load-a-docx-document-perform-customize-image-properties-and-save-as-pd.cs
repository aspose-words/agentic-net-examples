using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("Input.docx");

        // Create PDF save options and customize how images are rendered.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render all images in grayscale.
            ColorMode = ColorMode.Grayscale,
            // Use high‑quality rendering for better visual fidelity.
            UseHighQualityRendering = true
        };

        // Save the document as a PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
