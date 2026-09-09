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

        // Add seven pages with simple text.
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as a PDF file.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(pdfPath);

        // Pages to export (1‑based numbers).
        int[] pagesToExport = { 1, 4, 7 };
        // Desired resolution in DPI.
        const float resolutionDpi = 300f;

        foreach (int pageNumber in pagesToExport)
        {
            // Convert to zero‑based index for PageSet.
            int pageIndex = pageNumber - 1;

            // Configure image save options for PNG with custom resolution.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = resolutionDpi,
                PageSet = new PageSet(pageIndex)
            };

            string outFile = $"page_{pageNumber}.png";

            // Save the selected page as a PNG image.
            pdfDoc.Save(outFile, options);

            // Verify that the image file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create image file: {outFile}");
        }

        // All pages exported successfully.
        Console.WriteLine("Selected pages have been exported as PNG images.");
    }
}
