using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderAndSavePdf
{
    static void Main()
    {
        // Paths to the source document and the output PDF.
        string dataDir = @"C:\Data\";
        string artifactsDir = @"C:\Artifacts\";

        // Load the document that we want to render.
        Document doc = new Document(dataDir + "Rendering.docx");

        // Create a PdfSaveOptions object using the factory method.
        // This ensures we follow the provided lifecycle rules.
        PdfSaveOptions pdfOptions = (PdfSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Rendering options overview:
        // Enable anti‑aliasing for smoother edges.
        pdfOptions.UseAntiAliasing = true;
        // Use high‑quality (slow) rendering algorithms.
        pdfOptions.UseHighQualityRendering = true;
        // Optionally, control memory usage during saving.
        pdfOptions.MemoryOptimization = false;

        // Save the rendered document as a PDF.
        doc.Save(artifactsDir + "RenderedDocument.pdf", pdfOptions);
    }
}
