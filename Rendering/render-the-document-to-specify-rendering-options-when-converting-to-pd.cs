using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions object to specify rendering options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render the PDF with high‑quality (slower) algorithms.
            UseHighQualityRendering = true,

            // Enable anti‑aliasing for smoother graphics.
            UseAntiAliasing = true,

            // Render all colors in grayscale.
            ColorMode = ColorMode.Grayscale,

            // Embed all fonts fully into the PDF.
            EmbedFullFonts = true
        };

        // Save the document as a PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
