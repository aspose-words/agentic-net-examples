using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfToJpeg
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PDF document.
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is the first page of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page of the sample PDF.");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document for conversion.
        Document pdfDocument = new Document(pdfPath);

        // Configure image save options for a high‑quality JPEG with horizontal multi‑page layout.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 95,                     // High quality (0‑100).
            Resolution = 300,                     // Render at 300 dpi.
            UseHighQualityRendering = true,       // Enable high‑quality rendering algorithms.
            PageLayout = MultiPageLayout.Horizontal(10) // 10 pt spacing between pages.
        };

        // Save the PDF as a single JPEG image.
        string jpegPath = Path.Combine(outputDir, "result.jpg");
        pdfDocument.Save(jpegPath, jpegOptions);

        // Validate that the output file was created.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the JPEG image.");
        }

        Console.WriteLine($"PDF successfully exported to JPEG: {jpegPath}");
    }
}
