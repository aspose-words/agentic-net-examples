using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports OpenType ligatures (e.g., Calibri).
        builder.Font.Name = "Calibri";
        builder.Font.Size = 48;

        // Text containing common ligatures (fi, fl) and a stylistic set example.
        builder.Writeln("Office: fi, fl, ffi, ffl");
        builder.Writeln("Stylish text with ligatures.");

        // Configure FontSettings to point to the system fonts folder.
        // This ensures the required font is available during rendering.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, true);
        doc.FontSettings = fontSettings;

        // Set up image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages into a single multi‑page TIFF.
            PageSet = new PageSet(0), // 0 means start from the first page; all pages will be saved.
            Resolution = 300,        // 300 DPI for good quality.
            UseAntiAliasing = true,
            UseHighQualityRendering = true
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "RenderedDocument.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF output file.");

        // Optionally, output the path for debugging (no user interaction required).
        Console.WriteLine($"TIFF file saved to: {tiffPath}");
    }
}
