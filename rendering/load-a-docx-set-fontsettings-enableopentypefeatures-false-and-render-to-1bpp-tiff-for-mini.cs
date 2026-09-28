using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample DOCX and the rendered TIFF.
        string sourcePath = "sample.docx";
        string outputPath = "output.tiff";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document locally.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for TIFF rendering.");
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and configure FontSettings.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        FontSettings fontSettings = new FontSettings();
        // No OpenType feature toggle is required; we simply assign the settings.
        loadedDoc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 3. Render the document to a 1bpp (black‑and‑white) TIFF.
        // -----------------------------------------------------------------
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render as black‑and‑white (1 bit per pixel).
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Use CCITT Group 4 compression, suitable for 1bpp TIFF.
            TiffCompression = TiffCompression.Ccitt4
            // By default all pages are rendered; no explicit PageSet needed.
        };

        loadedDoc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // 4. Verify that the TIFF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF output file was not created.", outputPath);
        }

        Console.WriteLine($"Document rendered to 1bpp TIFF successfully: {Path.GetFullPath(outputPath)}");
    }
}
