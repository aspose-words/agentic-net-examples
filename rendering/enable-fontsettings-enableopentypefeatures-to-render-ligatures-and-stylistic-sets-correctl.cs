using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define an output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that supports ligatures (e.g., Arial) and write text containing the "fi" ligature.
        builder.Font.Name = "Arial";
        builder.Writeln("Office"); // Contains the "fi" ligature.

        // Configure font settings. Here we point to the system fonts folder so Aspose.Words can locate the font.
        FontSettings fontSettings = new FontSettings();
        string systemFontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (!string.IsNullOrEmpty(systemFontsFolder) && Directory.Exists(systemFontsFolder))
        {
            fontSettings.SetFontsFolder(systemFontsFolder, true);
        }
        // Assign the font settings to the document.
        doc.FontSettings = fontSettings;

        // Prepare TIFF rendering options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the first page only.
            PageSet = new PageSet(0),
            // Set a reasonable resolution.
            Resolution = 300
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "RenderedDocument.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception("Failed to create the TIFF output file.");
        }

        // Indicate successful completion.
        Console.WriteLine("TIFF rendering completed successfully: " + tiffPath);
    }
}
