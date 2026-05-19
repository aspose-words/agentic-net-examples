using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document that will be saved as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF that will be converted to a high‑resolution JPEG image.");

        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for JPEG.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // High resolution (e.g., 300 DPI) for better quality.
            Resolution = 300f,
            // Set JPEG quality to 100 (low compression, highest quality).
            JpegQuality = 100
        };

        const string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, jpegOptions);

        // Validate that the JPEG file was created and contains data.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
        {
            throw new InvalidOperationException("The JPEG image was not created successfully.");
        }

        // Clean up temporary files (optional).
        File.Delete(pdfPath);
    }
}
