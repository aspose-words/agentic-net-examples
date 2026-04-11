using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportSelectedPdfPages
{
    public static void Main()
    {
        // Folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document and save it as PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Generate 7 pages with simple text.
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Export selected pages (1, 4, 7) as separate PNG images.
        // -----------------------------------------------------------------
        int[] pagesToExport = { 1, 4, 7 }; // 1‑based page numbers.
        const float customResolution = 300f; // DPI.

        foreach (int pageNumber in pagesToExport)
        {
            // ImageSaveOptions controls raster image rendering.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Set the desired resolution (DPI).
                Resolution = customResolution,
                // Select the exact page (zero‑based index) to render.
                PageSet = new PageSet(pageNumber - 1)
            };

            string outFile = Path.Combine(artifactsDir, $"Page_{pageNumber}.png");
            pdfDoc.Save(outFile, options);

            // Validation: ensure the image file was created.
            if (!File.Exists(outFile) || new FileInfo(outFile).Length == 0)
                throw new InvalidOperationException($"Failed to create image for page {pageNumber}.");
        }

        // All done – the PNG files are located in the Artifacts folder.
    }
}
