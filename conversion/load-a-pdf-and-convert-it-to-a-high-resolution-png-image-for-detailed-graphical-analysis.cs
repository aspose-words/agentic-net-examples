using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file to act as the input.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        string inputPdfPath = "sample.pdf";
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("The input PDF file was not created.");

        // Load the PDF document.
        Document pdfDocument = new Document(inputPdfPath);

        // Configure high‑resolution PNG export (e.g., 300 DPI).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300
        };

        // Save the PDF as a PNG image.
        string outputPngPath = "output.png";
        pdfDocument.Save(outputPngPath, pngOptions);

        // Validate that the PNG file was created and has content.
        if (!File.Exists(outputPngPath) || new FileInfo(outputPngPath).Length == 0)
            throw new InvalidOperationException("The PNG image was not created or is empty.");

        // Optionally, clean up the temporary PDF file.
        // File.Delete(inputPdfPath);
    }
}
