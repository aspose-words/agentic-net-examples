using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfExtractor
{
    static void Main()
    {
        // Paths to the source PDF and the destination PDF.
        string inputPath = "input.pdf";
        string outputPath = "output.pdf";

        // Load the PDF document with default load options.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document sourceDoc = new Document(inputPath, loadOptions);

        // Extract the first page (page index is zero‑based, count is the number of pages to extract).
        Document extractedDoc = sourceDoc.ExtractPages(0, 1);

        // Configure PDF save options as needed.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Example: set PDF compliance level.
            Compliance = PdfCompliance.Pdf17
        };

        // Save the extracted content as a new PDF file.
        extractedDoc.Save(outputPath, saveOptions);
    }
}
