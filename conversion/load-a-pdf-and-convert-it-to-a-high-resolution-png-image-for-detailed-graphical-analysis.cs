using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be saved as PDF and then converted to a high‑resolution PNG.");

        // Save the document as PDF (input for conversion).
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF file.
        Document pdfDoc = new Document(pdfPath);

        // Configure high‑resolution PNG export options.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set resolution to 300 DPI for detailed analysis.
            Resolution = 300
        };

        // Save the PDF pages as PNG images.
        // Each page will be saved as a separate PNG file with a page number suffix.
        string pngBasePath = "output_page";
        pdfDoc.Save(pngBasePath + ".png", pngOptions);

        // Verify that at least one PNG file was created.
        string[] pngFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), pngBasePath + "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were created from the PDF.");

        // Optional: Ensure the first PNG file has content.
        FileInfo firstPng = new FileInfo(pngFiles[0]);
        if (firstPng.Length == 0)
            throw new InvalidOperationException("The generated PNG file is empty.");

        // Cleanup (optional): delete temporary files if desired.
        // File.Delete(pdfPath);
    }
}
