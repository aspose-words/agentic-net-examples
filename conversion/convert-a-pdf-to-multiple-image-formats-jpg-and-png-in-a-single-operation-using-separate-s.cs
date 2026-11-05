using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content generated for conversion.");
        source.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Convert PDF to JPEG.
        string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, SaveFormat.Jpeg);

        // Convert PDF to PNG.
        string pngPath = "output.png";
        pdfDoc.Save(pngPath, SaveFormat.Png);

        // Validate that the JPEG file was created.
        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("Expected JPEG output was not created.");

        // Validate that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("Expected PNG output was not created.");
    }
}
