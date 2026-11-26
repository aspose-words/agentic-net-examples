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
        builder.Writeln("Sample PDF content for archival purposes.");
        sourceDoc.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("sample.pdf");

        // Prepare output directory for PNG images.
        string outputDir = "output_images";
        Directory.CreateDirectory(outputDir);

        // Configure image save options for PNG (lossless by default).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Save each page of the PDF as a separate PNG file.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Save only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            string outputPath = Path.Combine(outputDir, $"page_{pageIndex + 1}.png");
            pdfDoc.Save(outputPath, pngOptions);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected PNG file was not created: {outputPath}");
        }

        // Verify that at least one PNG file was generated.
        string[] generatedFiles = Directory.GetFiles(outputDir, "*.png");
        if (generatedFiles.Length == 0)
            throw new InvalidOperationException("No PNG images were generated from the PDF.");

        // Example completed successfully.
    }
}
