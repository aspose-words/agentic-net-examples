using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document for image conversion.");

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert PDF to JPEG.
        string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, SaveFormat.Jpeg);

        // Convert PDF to PNG.
        string pngPath = "output.png";
        pdfDoc.Save(pngPath, SaveFormat.Png);

        // Validate that the image files were created.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("Expected JPEG output was not created.");

        if (!File.Exists(pngPath))
            throw new InvalidOperationException("Expected PNG output was not created.");

        Console.WriteLine("PDF successfully converted to JPEG and PNG.");
    }
}
