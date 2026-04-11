using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToPngConverter
{
    public static void Main()
    {
        // Define the folder where the PDF and PNG files will be stored.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Create a sample PDF document with three pages.
        string pdfPath = Path.Combine(outputFolder, "sample.pdf");
        CreateSamplePdf(pdfPath);

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Convert each page of the PDF to a separate PNG image.
        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex),

                // Optional: set resolution (dpi) for higher quality.
                Resolution = 300
            };

            string pngPath = Path.Combine(outputFolder, $"page_{pageIndex + 1}.png");
            pdfDocument.Save(pngPath, options);

            // Verify that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create image: {pngPath}");
        }

        // All pages have been converted successfully.
        Console.WriteLine($"PDF converted to PNG images in folder: {outputFolder}");
    }

    // Helper method to create a simple multi‑page PDF.
    private static void CreateSamplePdf(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        doc.Save(filePath, SaveFormat.Pdf);
    }
}
