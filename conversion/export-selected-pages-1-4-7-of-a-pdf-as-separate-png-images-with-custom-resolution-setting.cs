using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfPagesToPng
{
    public static void Main()
    {
        // Define file names and paths.
        const string pdfPath = "sample.pdf";
        const string outputFolder = "output";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document and save it as PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add at least 7 pages with simple text.
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i} of the sample PDF.");
            if (i < 7) // No break after the last page.
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create the source PDF at '{pdfPath}'.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and export selected pages as PNG images.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Pages to export (1‑based numbers).
        int[] pagesToExport = { 1, 4, 7 };
        // Desired resolution in DPI.
        const float resolutionDpi = 300f;

        foreach (int pageNumber in pagesToExport)
        {
            // Convert to zero‑based index for PageSet.
            int zeroBasedIndex = pageNumber - 1;

            // Configure image save options.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = resolutionDpi,
                PageSet = new PageSet(zeroBasedIndex)
            };

            // Build the output file name.
            string outFile = Path.Combine(outputFolder, $"Page_{pageNumber}.png");

            // Save the selected page as PNG.
            pdfDoc.Save(outFile, options);

            // Validate that the image was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create PNG for page {pageNumber} at '{outFile}'.");
        }

        // All done. The program exits automatically.
    }
}
