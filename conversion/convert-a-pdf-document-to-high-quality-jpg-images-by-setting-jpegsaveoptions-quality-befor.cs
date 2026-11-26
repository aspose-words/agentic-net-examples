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
        builder.Writeln("Sample PDF content - Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF content - Page 2.");

        // Save the document as PDF.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a high‑quality JPEG image.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 100, // Set maximum quality.
                PageSet = new PageSet(pageIndex) // Convert only the current page.
            };

            string jpegPath = $"page_{pageIndex + 1}.jpg";
            pdfDoc.Save(jpegPath, jpegOptions);

            if (!File.Exists(jpegPath))
                throw new InvalidOperationException($"Failed to create JPEG image for page {pageIndex + 1}.");
        }

        // Optional cleanup of the temporary PDF.
        // File.Delete(pdfPath);
    }
}
