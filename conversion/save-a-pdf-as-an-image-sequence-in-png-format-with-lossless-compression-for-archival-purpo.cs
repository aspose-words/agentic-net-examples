using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare directories
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a sample PDF document
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF page 3.");

        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF.");

        // Step 2: Load the PDF and render each page to a PNG image
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for PNG (lossless)
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // PNG is lossless by default; no additional compression settings required.

        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Render only the current page
            pngOptions.PageSet = new PageSet(pageIndex);

            string imagePath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.png");
            pdfDoc.Save(imagePath, pngOptions);

            // Validate that the image was created
            if (!File.Exists(imagePath))
                throw new InvalidOperationException($"Failed to create image for page {pageIndex + 1}.");
        }

        // Optional: Clean up the temporary PDF file
        File.Delete(pdfPath);
    }
}
