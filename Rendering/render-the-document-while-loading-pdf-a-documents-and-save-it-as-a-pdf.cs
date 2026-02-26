using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input PDF/A file path
        string inputPath = @"C:\Docs\InputPdfA.pdf";

        // Output PDF file path
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the PDF/A document with default load options
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        // Ensure images are not skipped during loading
        loadOptions.SkipPdfImages = false;

        // Load the document
        Document doc = new Document(inputPath, loadOptions);

        // Rebuild the page layout to guarantee correct rendering
        doc.UpdatePageLayout();

        // Configure PDF save options (default PDF format)
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Example: to save as PDF/A-2u uncomment the following line
        // saveOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the rendered document as PDF
        doc.Save(outputPath, saveOptions);
    }
}
