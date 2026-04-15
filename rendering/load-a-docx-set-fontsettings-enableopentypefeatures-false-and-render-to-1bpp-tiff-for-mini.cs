using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for rendering to 1bpp TIFF.");

        // Configure FontSettings (cannot set EnableOpenTypeFeatures due to restrictions).
        FontSettings fontSettings = new FontSettings();
        // Example: set a custom fonts folder if needed.
        // fontSettings.SetFontsFolder(@"C:\Windows\Fonts", false);
        doc.FontSettings = fontSettings;

        // Save the document to a temporary DOCX file.
        string docPath = "Sample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Load the document from the file.
        Document loadedDoc = new Document(docPath);

        // Set up ImageSaveOptions for 1bpp TIFF with minimal size.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 4 compression.
            TiffCompression = TiffCompression.Ccitt4,
            // Force 1-bit per pixel format.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Disable anti-aliasing and high‑quality rendering for smaller output.
            UseAntiAliasing = false,
            UseHighQualityRendering = false,
            // Optional resolution setting.
            Resolution = 300
        };

        // Render the document to TIFF.
        string tiffPath = "Output.tiff";
        loadedDoc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at {tiffPath}");
        }

        // Cleanup temporary DOCX file (optional).
        // File.Delete(docPath);
    }
}
