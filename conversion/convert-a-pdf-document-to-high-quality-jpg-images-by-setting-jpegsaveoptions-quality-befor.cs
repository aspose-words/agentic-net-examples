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
        builder.Writeln("This is a sample PDF document generated for conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a high‑quality JPG image.
        int pageCount = pdfDoc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            // Set JPEG quality to the maximum (100) using ImageSaveOptions.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            jpegOptions.JpegQuality = 100;
            jpegOptions.PageSet = new PageSet(i);

            string jpgPath = $"page_{i + 1}.jpg";
            pdfDoc.Save(jpgPath, jpegOptions);

            // Verify that the JPG image was created.
            if (!File.Exists(jpgPath))
                throw new InvalidOperationException($"The JPG file '{jpgPath}' was not created.");
        }
    }
}
