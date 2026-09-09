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
        const string outputFolder = "PngPages";

        // Ensure the output folder exists.
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample document and save it as PDF (input for conversion).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Export each page of the PDF to a separate PNG image at 300 DPI.
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options for PNG with 300 DPI.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300f,               // Set both horizontal and vertical DPI.
                PageSet = new PageSet(pageIndex) // Render only the current page.
            };

            string pngPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            pdfDoc.Save(pngPath, options);

            // Validate that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // -----------------------------------------------------------------
        // 4. Clean up temporary files (optional).
        // -----------------------------------------------------------------
        // File.Delete(pdfPath); // Uncomment if you want to remove the intermediate PDF.

        Console.WriteLine("PDF successfully exported to PNG images at 300 DPI.");
    }
}
