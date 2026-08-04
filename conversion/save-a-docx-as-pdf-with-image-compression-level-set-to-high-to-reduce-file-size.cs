using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample document with an image.");

        // Insert a tiny PNG image from a Base64 string (1x1 pixel).
        byte[] imageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        builder.InsertImage(imageBytes);

        string docxPath = "sample.docx";
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure PDF save options for high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg, // Compress all images as JPEG.
            JpegQuality = 10 // Low quality = high compression.
        };

        string pdfPath = "output.pdf";
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
