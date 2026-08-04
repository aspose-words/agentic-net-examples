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

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports ligatures (e.g., Calibri) and write text containing ligatures.
        builder.Font.Name = "Calibri";
        builder.Writeln("Office"); // Contains the 'fi' ligature.

        // Configure font settings (optional: point to a folder with additional fonts).
        FontSettings fontSettings = new FontSettings();
        // Example: set the system fonts folder (adjust if necessary for the environment).
        string systemFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
        if (Directory.Exists(systemFonts))
        {
            fontSettings.SetFontsFolder(systemFonts, true);
        }
        doc.FontSettings = fontSettings;

        // Set up image save options for TIFF rendering.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            UseAntiAliasing = true,
            UseHighQualityRendering = true,
            Resolution = 300 // DPI
        };

        // Render the document to a TIFF file.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception("TIFF rendering failed: output file not found.");
        }

        // Indicate successful completion.
        Console.WriteLine($"Document rendered to TIFF successfully: {tiffPath}");
    }
}
