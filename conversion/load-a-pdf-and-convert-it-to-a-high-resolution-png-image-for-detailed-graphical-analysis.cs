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
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to a high‑resolution PNG image.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for high resolution (e.g., 300 DPI).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300f // Sets both horizontal and vertical DPI.
        };

        // Save the first page of the PDF as a PNG image.
        pdfDoc.Save(pngPath, pngOptions);

        // -----------------------------------------------------------------
        // 3. Validate that the PNG image was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG image was not created.");

        // Optional: clean up the temporary PDF file.
        // File.Delete(pdfPath);
    }
}
