using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToJpegConverter
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Configure JPEG save options with high quality (low compression) and high resolution.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // JpegQuality ranges from 0 (worst) to 100 (best). 100 gives the lowest compression.
            JpegQuality = 100,
            // Set the resolution in DPI.
            Resolution = 300f
        };

        // Save the PDF as a high‑resolution JPEG image.
        string jpegPath = "output.jpeg";
        pdfDocument.Save(jpegPath, jpegOptions);

        // Verify that the JPEG file was created.
        if (!File.Exists(jpegPath))
        {
            throw new InvalidOperationException("The JPEG image was not created as expected.");
        }

        // Optional: Clean up temporary PDF file.
        // File.Delete(pdfPath);
    }
}
