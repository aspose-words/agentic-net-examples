using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document with two pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF page 2.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document we just created.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a high‑quality JPEG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            // Configure image save options for JPEG with maximum quality.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                PageSet = new PageSet(i), // Render only the current page.
                JpegQuality = 100          // Highest quality (least compression).
            };

            string imagePath = $"page_{i + 1}.jpg";
            pdfDoc.Save(imagePath, options);

            // Verify that the image file was created.
            if (!File.Exists(imagePath))
                throw new InvalidOperationException($"Failed to create image file: {imagePath}");
        }

        // Clean up temporary PDF file (optional).
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
