using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for image sequence conversion.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("sample.pdf");

        // Convert each page of the PDF to a separate PNG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.PageSet = new PageSet(i);               // Render the specific page.
            options.ImageColorMode = ImageColorMode.None;   // Preserve original colors (lossless PNG).
            options.Resolution = 300;                       // Optional: set DPI for high quality.

            string outputFileName = $"page_{i + 1}.png";
            pdfDoc.Save(outputFileName, options);

            // Verify that the image file was created.
            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Expected image file '{outputFileName}' was not created.");
        }

        // Verify that the PDF file exists.
        if (!File.Exists("sample.pdf"))
            throw new InvalidOperationException("Expected PDF file 'sample.pdf' was not created.");
    }
}
