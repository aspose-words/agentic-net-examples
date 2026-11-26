using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for image conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Save as JPEG.
        string jpegPath = "output.jpg";
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        pdfDoc.Save(jpegPath, jpegOptions);

        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("JPEG image was not created.");

        // Save as PNG.
        string pngPath = "output.png";
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pdfDoc.Save(pngPath, pngOptions);

        if (!File.Exists(pngPath))
            throw new InvalidOperationException("PNG image was not created.");

        // Indicate success (optional, not interactive).
        Console.WriteLine("PDF successfully converted to JPEG and PNG.");
    }
}
