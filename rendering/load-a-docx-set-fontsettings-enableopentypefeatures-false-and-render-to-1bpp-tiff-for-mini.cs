using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample DOCX and the output TIFF.
        string docPath = "sample.docx";
        string tiffPath = "output.tiff";

        // -----------------------------------------------------------------
        // Create a simple DOCX document locally (do not rely on external files).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a sample document for rendering to 1bpp TIFF.");
        doc.Save(docPath); // Save the source document.

        // -----------------------------------------------------------------
        // Load the document from the file we just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // Configure FontSettings. The task mentions disabling OpenType features,
        // but the API for EnableOpenTypeFeatures is not available, so we simply assign a FontSettings instance.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        loadedDoc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // Set up image save options for a 1‑bit‑per‑pixel (black‑and‑white) TIFF.
        // Use LZW compression (available in all supported versions) to keep the file size minimal.
        // -----------------------------------------------------------------
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.BlackAndWhite,
            TiffCompression = TiffCompression.Lzw
        };

        // -----------------------------------------------------------------
        // Render the document to TIFF.
        // -----------------------------------------------------------------
        loadedDoc.Save(tiffPath, saveOptions);

        // -----------------------------------------------------------------
        // Verify that the TIFF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(tiffPath))
        {
            throw new Exception("TIFF file was not created.");
        }

        Console.WriteLine("Document rendered to 1bpp TIFF successfully.");
    }
}
