using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample multi‑page PDF.
        const string pdfPath = "sample.pdf";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 10)
                builder.InsertBreak(BreakType.PageBreak);
        }
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF.");

        // Load the PDF for conversion.
        Document pdfDoc = new Document(pdfPath);

        // Pages to export as PNG.
        int[] pagesToExport = { 1, 4, 7 };
        const int resolution = 300; // DPI

        foreach (int pageNumber in pagesToExport)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageNumber),
                HorizontalResolution = resolution,
                VerticalResolution = resolution
            };

            string outputPath = $"page_{pageNumber}.png";
            pdfDoc.Save(outputPath, options);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"PNG for page {pageNumber} was not created.");
        }

        // Optional cleanup: delete the temporary PDF if not needed.
        // File.Delete(pdfPath);
    }
}
