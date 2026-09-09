using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("First page of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the sample PDF.");

        // Save the sample as PDF.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a high‑quality JPEG image.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Highest quality (0‑100). 100 means no compression loss.
                JpegQuality = 100,
                // Use slower but higher‑quality rendering algorithms.
                UseHighQualityRendering = true,
                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };

            string jpegPath = $"output_page_{pageIndex + 1}.jpg";
            pdfDoc.Save(jpegPath, options);

            // Verify that the image was created.
            if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
                throw new InvalidOperationException($"Failed to create JPEG image for page {pageIndex + 1}.");
        }

        Console.WriteLine("PDF successfully converted to high‑quality JPEG images.");
    }
}
