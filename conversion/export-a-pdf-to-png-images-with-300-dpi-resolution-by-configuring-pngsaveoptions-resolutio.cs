using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for PNG export.");
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Export each page of the PDF to a PNG image with 300 DPI resolution.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            string outputPath = $"page_{pageIndex + 1}.png";

            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300,
                PageSet = new PageSet(pageIndex) // Export only the current page.
            };

            pdfDoc.Save(outputPath, pngOptions);

            // Validate that the PNG file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output PNG '{outputPath}' was not created.");
        }

        // Optional: confirm that at least one PNG was generated.
        if (pdfDoc.PageCount == 0)
            throw new InvalidOperationException("The PDF document contains no pages to export.");
    }
}
