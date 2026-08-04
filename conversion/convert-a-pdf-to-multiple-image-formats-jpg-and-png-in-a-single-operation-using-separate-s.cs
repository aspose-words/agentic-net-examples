using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for image conversion.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the generated PDF.
        Document pdfDoc = new Document("sample.pdf");

        // Convert the PDF to JPEG image.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        pdfDoc.Save("sample.jpg", jpegOptions);

        // Convert the PDF to PNG image.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pdfDoc.Save("sample.png", pngOptions);

        // Validate that the output files were created.
        if (!File.Exists("sample.jpg"))
            throw new InvalidOperationException("JPEG image was not created.");

        if (!File.Exists("sample.png"))
            throw new InvalidOperationException("PNG image was not created.");
    }
}
