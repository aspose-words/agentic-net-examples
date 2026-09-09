using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string pdfPath = "sample.pdf";
        const string jpegPath = "output.jpg";

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1: Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2: Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3: Ut enim ad minim veniam, quis nostrud exercitation ullamco.");

        // Save the document as PDF (bootstrap input file).
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF that we just created.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure image save options for a high‑quality JPEG.
        // -----------------------------------------------------------------
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Highest JPEG quality (0‑100).
            JpegQuality = 100,

            // Render all pages side by side in a single image.
            PageLayout = MultiPageLayout.Horizontal(10f),

            // Optional: improve rendering quality.
            UseAntiAliasing = true,
            UseHighQualityRendering = true
        };

        // Save the PDF as a single JPEG image.
        pdfDoc.Save(jpegPath, options);

        // -----------------------------------------------------------------
        // 4. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
        {
            throw new InvalidOperationException("The JPEG image was not created successfully.");
        }

        // Cleanup temporary files (optional).
        // File.Delete(pdfPath);
        // File.Delete(jpegPath);
    }
}
