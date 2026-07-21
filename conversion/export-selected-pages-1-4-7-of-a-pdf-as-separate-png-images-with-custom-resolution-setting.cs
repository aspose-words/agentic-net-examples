using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfPagesToPng
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string outputDir = "output";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document and save it as PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add seven pages with simple text.
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and export selected pages as PNG images.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Pages to export: 1, 4, 7 (zero‑based indices 0, 3, 6).
        int[] pageIndices = { 0, 3, 6 };
        const float dpi = 300f; // Custom resolution.

        foreach (int pageIndex in pageIndices)
        {
            // Configure image save options.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = dpi,               // Set both horizontal and vertical DPI.
                PageSet = new PageSet(pageIndex) // Render only the specified page.
            };

            // Build the output file name.
            string outFile = Path.Combine(outputDir, $"Page{pageIndex + 1}.png");

            // Save the selected page as a PNG image.
            pdfDoc.Save(outFile, options);

            // Validate that the image was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Image for page {pageIndex + 1} was not created.");
        }

        // All done – the program exits automatically.
    }
}
