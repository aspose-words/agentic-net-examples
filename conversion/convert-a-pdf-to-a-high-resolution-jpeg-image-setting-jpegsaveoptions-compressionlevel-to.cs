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
        builder.Writeln("Sample PDF content for conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for high‑resolution JPEG with low compression (high quality).
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // JpegQuality ranges from 0 (lowest) to 100 (highest). Use a high value for low compression.
            JpegQuality = 100,
            // Set the resolution in DPI.
            Resolution = 300
        };

        // Save the PDF as a JPEG image.
        string jpegPath = "output.jpeg";
        pdfDoc.Save(jpegPath, jpegOptions);

        // Validate that the JPEG file was created.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("Expected output JPEG was not created.");
    }
}
