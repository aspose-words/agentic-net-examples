using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Configure PDF rendering options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Enable high‑quality (slow) rendering algorithms.
            UseHighQualityRendering = true,

            // Enable anti‑aliasing for smoother edges.
            UseAntiAliasing = true,

            // Render DrawingML shapes themselves (not fall‑backs).
            DmlRenderingMode = DmlRenderingMode.DrawingML,

            // Render DrawingML effects with the highest quality.
            DmlEffectsRenderingMode = DmlEffectsRenderingMode.Fine,

            // Render colors in normal mode (full color).
            ColorMode = ColorMode.Normal,

            // Configure metafile rendering to use vector rendering with bitmap fallback.
            MetafileRenderingOptions = new MetafileRenderingOptions
            {
                EmulateRasterOperations = false,
                RenderingMode = MetafileRenderingMode.VectorWithFallback
            }
        };

        // Save the document as PDF using the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
