using System.Drawing.Text;                     // For TextRenderingHint enum
using Aspose.Words;                               // Core document API
using Aspose.Words.Saving;                        // PDF save options

// Load an existing Word document.
Document doc = new Document("Input.docx");

// Create PDF save options and configure text rendering.
// Enable anti‑aliasing and high‑quality rendering for sharper text.
// If a specific TextRenderingHint is required, assign it via
// the GraphicsQualityOptions property (available on FixedPageSaveOptions).
PdfSaveOptions pdfOptions = new PdfSaveOptions
{
    UseAntiAliasing = true,
    UseHighQualityRendering = true,
    // Uncomment the following lines if you need to set a precise text rendering hint.
    // GraphicsQualityOptions = new GraphicsQualityOptions
    // {
    //     TextRenderingHint = TextRenderingHint.ClearTypeGridFit
    // }
};

// Save the document as PDF using the configured options.
doc.Save("Output.pdf", pdfOptions);
