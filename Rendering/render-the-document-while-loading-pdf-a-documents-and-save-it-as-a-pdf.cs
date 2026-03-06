using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF/A document.
        string inputPath = "InputPdfA.pdf";

        // Path where the rendered PDF will be saved.
        string outputPath = "RenderedOutput.pdf";

        // Load the PDF/A document using PdfLoadOptions.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Rebuild the page layout to ensure correct rendering.
        doc.UpdatePageLayout();

        // Create save options appropriate for PDF format.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Configure PDF save options (optional high‑quality rendering).
        if (saveOptions is PdfSaveOptions pdfOptions)
        {
            pdfOptions.UseHighQualityRendering = true;
            // Save as a regular PDF (PDF 1.7 compliance).
            pdfOptions.Compliance = PdfCompliance.Pdf17;
        }

        // Save the rendered document as PDF.
        doc.Save(outputPath, saveOptions);
    }
}
