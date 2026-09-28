using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample multi‑page document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is the content of page {i}.");
            if (i < 7)
            {
                // Insert a page break to start a new page.
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as PDF (input for the conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The source PDF was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Pages to export (1‑based page numbers).
        int[] pagesToExport = { 1, 4, 7 };
        const int resolutionDpi = 300; // Custom resolution.

        foreach (int pageNumber in pagesToExport)
        {
            // Aspose.Words uses zero‑based page indexes.
            int pageIndex = pageNumber - 1;

            // Configure image save options for PNG with custom resolution.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex),
                Resolution = resolutionDpi
            };

            string outputPath = $"page_{pageNumber}.png";
            pdfDoc.Save(outputPath, options);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output image '{outputPath}' was not created.");
        }
    }
}
