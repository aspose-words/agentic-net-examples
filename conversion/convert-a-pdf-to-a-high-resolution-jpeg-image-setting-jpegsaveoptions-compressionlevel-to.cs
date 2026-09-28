using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document generated for conversion.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The source PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure JPEG save options: high resolution and low compression (high quality).
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // JpegQuality ranges from 0 (worst) to 100 (best). 100 gives lowest compression.
            JpegQuality = 100,
            Resolution = 300 // DPI for high‑resolution output.
        };

        const string jpegPath = "output.jpeg";
        pdfDoc.Save(jpegPath, jpegOptions);

        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("The JPEG image was not created.");

        // Optional: clean up temporary PDF file.
        try
        {
            File.Delete(pdfPath);
        }
        catch
        {
            // Ignore any errors during cleanup.
        }
    }
}
