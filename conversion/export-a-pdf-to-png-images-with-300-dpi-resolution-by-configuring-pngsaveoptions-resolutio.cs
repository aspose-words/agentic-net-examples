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
        builder.Writeln("This is a sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is a sample PDF page 2.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("sample.pdf");

        // Configure image save options for PNG with 300 DPI resolution.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300
        };

        // Export each page of the PDF to a separate PNG file.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Export only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);
            string outputFileName = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(outputFileName, pngOptions);

            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Expected PNG file '{outputFileName}' was not created.");
        }

        // Verify that at least one PNG file was created.
        bool anyPngCreated = false;
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            if (File.Exists($"page_{i + 1}.png"))
            {
                anyPngCreated = true;
                break;
            }
        }

        if (!anyPngCreated)
            throw new InvalidOperationException("No PNG files were created from the PDF document.");

        // Optional cleanup (commented out).
        // File.Delete("sample.pdf");
    }
}
