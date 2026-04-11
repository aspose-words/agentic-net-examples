using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for temporary files and output images.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string outputDir = Path.Combine(artifactsDir, "PngPages");

        // Ensure the directories exist.
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add three pages with simple text.
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Save the document as PDF.
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // Step 3: Export each page of the PDF as a separate PNG image.
        // -----------------------------------------------------------------
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(i)
            };

            string pngPath = Path.Combine(outputDir, $"Page_{i + 1}.png");

            // Save the current page as PNG.
            pdfDoc.Save(pngPath, options);

            // Validate that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to save PNG for page {i + 1} at '{pngPath}'.");
        }

        // All pages have been exported successfully.
        Console.WriteLine($"Exported {pdfDoc.PageCount} PNG images to '{outputDir}'.");
    }
}
