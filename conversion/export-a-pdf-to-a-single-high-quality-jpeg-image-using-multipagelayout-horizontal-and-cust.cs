using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate PDF and final JPEG.
        const string pdfPath = "sample.pdf";
        const string jpegPath = "output.jpg";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample PDF content for JPEG conversion.");
        builder.Writeln("This document will be rendered as a single high‑quality JPEG image.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and export it to a single JPEG image.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Render all pages side‑by‑side horizontally with a 10‑point spacing.
            PageLayout = MultiPageLayout.Horizontal(10f),
            // Set JPEG quality to the maximum (100) for high quality.
            JpegQuality = 100,
            // Improve rendering quality.
            UseAntiAliasing = true,
            UseHighQualityRendering = true
        };

        // Save the PDF as a JPEG image.
        pdfDoc.Save(jpegPath, options);

        // Verify that the JPEG was created.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("JPEG image was not created or is empty.");
    }
}
