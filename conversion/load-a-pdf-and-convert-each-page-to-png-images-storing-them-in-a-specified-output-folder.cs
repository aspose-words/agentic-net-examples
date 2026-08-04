using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define the folder where the PDF and PNG images will be stored.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "PdfToPngOutput");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with multiple pages.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        string pdfPath = Path.Combine(outputFolder, "sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);

        // -----------------------------------------------------------------
        // Step 3: Convert each page of the PDF to a separate PNG image.
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Configure the options to render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            string pngPath = Path.Combine(outputFolder, $"page_{pageIndex + 1}.png");
            pdfDocument.Save(pngPath, pngOptions);

            // Validate that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // -----------------------------------------------------------------
        // Completion message (optional, not required for non‑interactive run).
        // -----------------------------------------------------------------
        Console.WriteLine($"PDF converted to PNG images successfully. Files are located in: {outputFolder}");
    }
}
