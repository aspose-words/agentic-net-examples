using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create PDF save options and configure rendering settings.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render all colors as grayscale.
            ColorMode = ColorMode.Grayscale,

            // Use high‑quality (slow) rendering algorithms.
            UseHighQualityRendering = true,

            // Enable anti‑aliasing for smoother edges.
            UseAntiAliasing = true,

            // Render DrawingML shapes themselves (not their fallback shapes).
            DmlRenderingMode = DmlRenderingMode.DrawingML,

            // Render DrawingML effects with the highest quality.
            DmlEffectsRenderingMode = DmlEffectsRenderingMode.Fine,

            // Set the initial zoom factor (percentage) when the PDF is opened.
            ZoomFactor = 100
        };

        // Show hidden text in the rendered PDF.
        doc.LayoutOptions.ShowHiddenText = true;

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
