using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document that will be saved as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document.");
        builder.Writeln("It will be converted to high‑quality JPG images.");

        // Save the document as PDF (input for conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a high‑quality JPG image.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options for JPEG with maximum quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 100,               // High quality (0‑100)
                PageSet = new PageSet(pageIndex) // Save only the current page.
            };

            string jpgPath = $"page_{pageIndex + 1}.jpg";
            pdfDoc.Save(jpgPath, jpegOptions);

            if (!File.Exists(jpgPath) || new FileInfo(jpgPath).Length == 0)
                throw new InvalidOperationException($"Failed to create JPG for page {pageIndex + 1}.");
        }

        // All pages have been converted successfully.
        Console.WriteLine("PDF conversion to high‑quality JPG images completed.");
    }
}
