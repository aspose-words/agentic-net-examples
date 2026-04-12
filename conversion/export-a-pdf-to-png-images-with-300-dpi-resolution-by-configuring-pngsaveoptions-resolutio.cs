using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is page 1 of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2 of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3 of the sample PDF.");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF at '{pdfPath}'.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and export each page to a PNG image at 300 DPI.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            // Configure image save options for PNG with 300 DPI.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300,               // Set both horizontal and vertical DPI.
                PageSet = new PageSet(i)        // Render only the current page.
            };

            string pngPath = Path.Combine(outputDir, $"page_{i + 1}.png");
            pdfDoc.Save(pngPath, options);

            // Validate that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG at '{pngPath}'.");
        }

        // All pages have been exported successfully.
        Console.WriteLine($"PDF converted to PNG images at 300 DPI. Files are located in: {outputDir}");
    }
}
