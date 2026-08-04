using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;          // Needed for FontSettings
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Create a sample DOCX document.
        string docPath = Path.Combine(dataDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a sample document for TIFF rendering.");
        doc.Save(docPath);

        // Load the document.
        Document loadedDoc = new Document(docPath);

        // Apply font settings (without using the banned EnableOpenTypeFeatures property).
        FontSettings fontSettings = new FontSettings();
        loadedDoc.FontSettings = fontSettings;

        // Configure image save options for 1bpp TIFF (minimal size).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // 1 bit per pixel.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Use CCITT Group 3 compression suitable for 1bpp images.
            TiffCompression = TiffCompression.Ccitt3,
            // Set resolution (dots per inch) for the rendered image.
            Resolution = 300
        };

        // Render the document to a TIFF file.
        string tiffPath = Path.Combine(dataDir, "output.tiff");
        loadedDoc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF output file.");

        Console.WriteLine($"TIFF rendered successfully: {tiffPath}");
    }
}
