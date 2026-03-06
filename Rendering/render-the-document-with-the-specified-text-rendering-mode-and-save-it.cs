// Load an existing Word document.
Aspose.Words.Document doc = new Aspose.Words.Document("InputDocument.docx");

// Create PDF save options to control rendering.
Aspose.Words.Saving.PdfSaveOptions pdfOptions = new Aspose.Words.Saving.PdfSaveOptions
{
    // Enable anti‑aliasing for smoother text rendering.
    UseAntiAliasing = true,

    // Use high‑quality rendering algorithms (slower but better visual fidelity).
    UseHighQualityRendering = true
};

// Save the document as PDF using the configured options.
doc.Save("RenderedDocument.pdf", pdfOptions);
