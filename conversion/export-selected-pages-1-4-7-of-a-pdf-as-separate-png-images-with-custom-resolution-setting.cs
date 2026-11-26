using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfPagesToPng
{
    public static void Main()
    {
        // Create a sample multi-page document.
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

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF.
        Document pdfDoc = new Document(pdfPath);

        // Pages to export (1‑based).
        int[] pagesToExport = { 1, 4, 7 };
        const int resolutionDpi = 300;

        foreach (int pageNumber in pagesToExport)
        {
            // Configure image save options.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = resolutionDpi,
                // PageSet uses zero‑based page index.
                PageSet = new PageSet(pageNumber - 1)
            };

            string outputPath = $"page_{pageNumber}.png";
            pdfDoc.Save(outputPath, options);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"PNG for page {pageNumber} was not created.");
        }
    }
}
