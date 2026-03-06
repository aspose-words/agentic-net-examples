using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF/UA document.
        string inputPath = @"C:\Docs\InputPdfUa.pdf";

        // Path where the rendered PDF will be saved.
        string outputPath = @"C:\Docs\RenderedOutput.pdf";

        // Load the PDF using PdfLoadOptions (default settings are sufficient for PDF/UA).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Rebuild the page layout to ensure accurate rendering.
        doc.UpdatePageLayout();

        // Configure save options for PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set compliance to PDF/UA‑1 (ISO 14289‑1).
            Compliance = PdfCompliance.PdfUa1,

            // Required flag for PDF/UA: display the document title in the viewer's title bar.
            DisplayDocTitle = true,

            // Export document structure (tags). This is mandatory for PDF/UA and will be applied automatically.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
