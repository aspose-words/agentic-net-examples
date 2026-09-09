using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample Word document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for high‑resolution PNG conversion.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Step 2: Load the generated PDF.
        Document pdfDoc = new Document(pdfPath);

        // Step 3: Configure image save options for high resolution (e.g., 300 DPI).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300 // Dots per inch.
        };

        // Step 4: Save the first page of the PDF as a PNG image.
        const string pngPath = "output.png";
        pdfDoc.Save(pngPath, pngOptions);

        // Step 5: Validate that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG image was not created as expected.");

        // Optional: Clean up intermediate PDF if not needed.
        // File.Delete(pdfPath);
    }
}
