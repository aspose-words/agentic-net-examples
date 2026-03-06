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
            // Enable anti‑aliasing for smoother edges.
            UseAntiAliasing = true,
            // Use high‑quality rendering (slower but better visual quality).
            UseHighQualityRendering = true,
            // Render DrawingML shapes directly instead of fallback shapes.
            DmlRenderingMode = DmlRenderingMode.DrawingML,
            // Preserve document structure for accessibility.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF using the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
