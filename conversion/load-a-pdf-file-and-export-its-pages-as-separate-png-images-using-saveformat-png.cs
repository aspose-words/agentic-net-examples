using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define an output folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with three pages.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and export each page as a separate PNG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Prepare ImageSaveOptions for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set a higher resolution for better quality.
            Resolution = 300
        };

        // Export each page.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure the page to render.
            pngOptions.PageSet = new PageSet(pageIndex);

            string pngPath = Path.Combine(outputDir, $"page_{pageIndex + 1}.png");
            pdfDoc.Save(pngPath, pngOptions);

            // Validate that the PNG file was written.
            if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                throw new InvalidOperationException($"PNG for page {pageIndex + 1} was not created.");
        }

        // All pages have been exported successfully.
        Console.WriteLine($"PDF converted to PNG images. Files are located in: {outputDir}");
    }
}
