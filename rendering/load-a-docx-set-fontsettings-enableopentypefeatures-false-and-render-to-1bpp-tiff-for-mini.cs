using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX document.
        string sampleDocPath = Path.Combine(outputDir, "Sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document used for rendering to a 1bpp TIFF image.");
        sampleDoc.Save(sampleDocPath);

        // Load the DOCX document.
        Document doc = new Document(sampleDocPath);

        // Configure FontSettings (without using banned OpenType APIs).
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Set up image save options for 1bpp TIFF with CCITT compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression which is suitable for 1bpp images.
            TiffCompression = TiffCompression.Ccitt4,
            // Force the pixel format to 1bpp indexed.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Disable anti-aliasing to keep the image binary.
            UseAntiAliasing = false,
            // Use default resolution (72 DPI) – can be adjusted if needed.
            Resolution = 300
        };

        // Render and save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "Rendered.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF output file.");

        // Optionally, output the result path.
        Console.WriteLine($"TIFF image saved to: {tiffPath}");
    }
}
