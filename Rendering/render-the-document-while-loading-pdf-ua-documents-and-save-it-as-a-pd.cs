using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class RenderPdfUaToPdf
{
    static void Main()
    {
        // Path to the source PDF/UA document.
        string inputPath = @"C:\Docs\SourcePdfUa.pdf";

        // Path where the rendered PDF will be saved.
        string outputPath = @"C:\Docs\RenderedOutput.pdf";

        // Load the PDF/UA document with default load options.
        // PdfLoadOptions can be customized if needed (e.g., SkipPdfImages, Password, etc.).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Ensure the document layout is up‑to‑date before rendering.
        doc.UpdatePageLayout();

        // Configure PDF save options for PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/UA‑1 (ISO 14289‑1). Use PdfUa2 for PDF/UA‑2 if required.
            Compliance = PdfCompliance.PdfUa1,

            // Required for PDF/UA: display the document title in the viewer's title bar.
            DisplayDocTitle = true,

            // Export the document structure (tags) – mandatory for PDF/UA.
            ExportDocumentStructure = true,

            // Optional: improve accessibility by preserving form fields.
            PreserveFormFields = true
        };

        // Save the document as a PDF using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
