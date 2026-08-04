using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfPagesToPng
{
    public static void Main()
    {
        // Create a sample multi‑page document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(pdfPath);

        // Pages to export (zero‑based indices).
        int[] pageIndices = { 0, 3, 6 };
        const float resolutionDpi = 300f; // Custom resolution.

        foreach (int pageIndex in pageIndices)
        {
            // Configure image save options for PNG with the desired resolution.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex),
                Resolution = resolutionDpi
            };

            string outFileName = $"Page_{pageIndex + 1}.png";
            pdfDoc.Save(outFileName, options);

            // Verify that the image file was created.
            if (!File.Exists(outFileName))
                throw new InvalidOperationException($"Failed to create image file: {outFileName}");
        }
    }
}
