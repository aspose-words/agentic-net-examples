using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths that simulate a network share and local output directories.
        string baseDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
        string networkShareDir = Path.Combine(baseDir, "NetworkShare");
        string localOutputDir = Path.Combine(baseDir, "Output");
        string customFontDir = Path.Combine(baseDir, "CustomFonts");

        // Ensure all directories exist.
        Directory.CreateDirectory(networkShareDir);
        Directory.CreateDirectory(localOutputDir);
        Directory.CreateDirectory(customFontDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document locally.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document created for rendering.");
        builder.Writeln("It will be saved to a simulated network share, loaded,");
        builder.Writeln("custom font settings will be applied, and finally rendered to TIFF.");

        // Save the sample document to the simulated network share.
        string networkDocPath = Path.Combine(networkShareDir, "SampleDocument.docx");
        sampleDoc.Save(networkDocPath);

        // -----------------------------------------------------------------
        // 2. Load the document from the network share.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(networkDocPath);

        // -----------------------------------------------------------------
        // 3. Configure custom font settings.
        //    (The folder may be empty; Aspose.Words will fall back to system fonts.)
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontDir, recursive: true);
        loadedDoc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 4. Render the document to a multipage TIFF image.
        // -----------------------------------------------------------------
        string tiffOutputPath = Path.Combine(localOutputDir, "RenderedDocument.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: set resolution (dpi) for higher quality.
            Resolution = 300
        };

        loadedDoc.Save(tiffOutputPath, tiffOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the TIFF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(tiffOutputPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{tiffOutputPath}'.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document rendered to TIFF successfully:");
        Console.WriteLine(tiffOutputPath);
    }
}
