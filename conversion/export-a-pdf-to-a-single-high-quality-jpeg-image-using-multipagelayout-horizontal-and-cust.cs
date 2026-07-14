using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfToJpeg
{
    public static void Main()
    {
        // Step 1: Create a sample multi‑page document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Page 1 - Sample content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 - More content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 - Final page.");

        // Step 2: Save the document as PDF (the source for conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The source PDF was not created.");

        // Step 3: Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Step 4: Configure image save options for a high‑quality JPEG.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 100,                 // Highest quality (0‑100).
            UseHighQualityRendering = true,    // Use high‑quality rendering algorithms.
            UseAntiAliasing = true,            // Enable anti‑aliasing.
            PageLayout = MultiPageLayout.Horizontal(10) // Render pages side‑by‑side with 10 pts spacing.
        };

        // Step 5: Save the PDF as a single JPEG image.
        const string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, jpegOptions);

        if (!File.Exists(jpegPath))
            throw new InvalidOperationException("The JPEG image was not created.");

        Console.WriteLine("PDF successfully exported to high‑quality JPEG: " + Path.GetFullPath(jpegPath));
    }
}
