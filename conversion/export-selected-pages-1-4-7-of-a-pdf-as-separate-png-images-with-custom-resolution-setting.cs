using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF with at least 7 pages.
        const string pdfPath = "sample.pdf";
        Document pdfDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDocument);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        pdfDocument.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF.");

        // Load the PDF we just created.
        Document loadedPdf = new Document(pdfPath);

        // Pages to export (1‑based page numbers: 1, 4, 7).
        int[] pageIndicesZeroBased = { 0, 3, 6 };
        const float customResolution = 300f; // DPI

        foreach (int pageIndex in pageIndicesZeroBased)
        {
            // Configure image save options for PNG with custom resolution.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex),
                Resolution = customResolution
            };

            string outputFileName = $"Page_{pageIndex + 1}.png";
            loadedPdf.Save(outputFileName, options);

            // Verify that the image was created.
            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Image for page {pageIndex + 1} was not created.");
        }
    }
}
