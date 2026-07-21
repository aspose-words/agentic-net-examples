using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfToJpeg
{
    public static void Main()
    {
        // Paths for temporary files.
        const string pdfPath = "sample.pdf";
        const string jpegPath = "output.jpg";

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document and save it as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages with simple text.
        builder.Writeln("Page 1: Hello World!");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2: Aspose.Words conversion example.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3: Exporting PDF to a single JPEG image.");

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and export it to a single high‑quality JPEG.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // High JPEG quality (0‑100). 100 = best quality, larger file size.
            JpegQuality = 100,

            // Render all pages side by side horizontally.
            PageLayout = MultiPageLayout.Horizontal(10f),

            // Optional: improve rendering quality.
            UseAntiAliasing = true,
            UseHighQualityRendering = true
        };

        // Save the rendered image.
        pdfDoc.Save(jpegPath, options);

        // -----------------------------------------------------------------
        // 3. Validate the output JPEG.
        // -----------------------------------------------------------------
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("The JPEG image was not created successfully.");

        // Indicate success (no interactive output required).
        Console.WriteLine("PDF successfully exported to JPEG: " + jpegPath);
    }
}
