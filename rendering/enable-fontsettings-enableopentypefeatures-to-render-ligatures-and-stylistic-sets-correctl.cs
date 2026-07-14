using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define output directory and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "Ligatures.tiff");

        // Create a new document and add text that contains common ligatures.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports ligatures (e.g., Arial).
        builder.Font.Name = "Arial";
        builder.Font.Size = 48;
        builder.Writeln("Ligatures demonstration:");
        builder.Writeln("fi fl ffi ffl"); // Text with ligatures.

        // Configure FontSettings to point to the system fonts folder.
        // This ensures the renderer can locate the required font files.
        FontSettings fontSettings = new FontSettings();
        string systemFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(systemFonts, true);
        doc.FontSettings = fontSettings;

        // Set up image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render with high quality to preserve typographic features.
            UseAntiAliasing = true,
            UseHighQualityRendering = true,
            // Optional: increase resolution for clearer output.
            Resolution = 300
        };

        // Render the document to a TIFF file.
        doc.Save(tiffPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // The example finishes without requiring user interaction.
    }
}
