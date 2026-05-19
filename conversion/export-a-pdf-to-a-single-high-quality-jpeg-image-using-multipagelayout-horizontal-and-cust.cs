using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF (input.pdf).
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("First page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page of the PDF.");
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document("input.pdf");

        // Configure image save options for a single high‑quality JPEG.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // High quality (0‑100). 100 = best quality, least compression.
            JpegQuality = 100,
            // Render all pages side by side horizontally.
            PageLayout = MultiPageLayout.Horizontal(10f), // 10 points spacing between pages.
            // Optional: improve rendering quality.
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };

        // Save the PDF as a single JPEG image.
        string outputPath = "output.jpg";
        pdfDoc.Save(outputPath, jpegOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("The JPEG image was not created successfully.");

        // Example completed without requiring user interaction.
    }
}
