using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string pngPath = "output.png";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for high resolution (e.g., 300 DPI).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300 // Sets both horizontal and vertical DPI.
        };

        // Save the first page of the PDF as a PNG image.
        pdfDoc.Save(pngPath, pngOptions);

        // -----------------------------------------------------------------
        // Step 3: Validate the PNG output.
        // -----------------------------------------------------------------
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG image was not created.");

        FileInfo pngInfo = new FileInfo(pngPath);
        if (pngInfo.Length == 0)
            throw new InvalidOperationException("The PNG image file is empty.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine("PDF successfully converted to high‑resolution PNG.");
    }
}
