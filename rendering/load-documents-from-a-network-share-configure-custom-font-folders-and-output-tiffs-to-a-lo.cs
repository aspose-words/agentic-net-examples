using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for all example files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExampleData");
        Directory.CreateDirectory(baseDir);

        // Simulated network share folder.
        string networkShareDir = Path.Combine(baseDir, "NetworkShare");
        Directory.CreateDirectory(networkShareDir);

        // Custom fonts folder.
        string customFontDir = Path.Combine(baseDir, "CustomFonts");
        Directory.CreateDirectory(customFontDir);

        // Output folder for TIFF files.
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple source document.
        Document sampleDoc = new Document();
        sampleDoc.FirstSection.Body.AppendParagraph("Hello from the network share!");
        string sourcePath = Path.Combine(networkShareDir, "sample.docx");
        sampleDoc.Save(sourcePath);

        // Load the document from the simulated network share.
        Document doc = new Document(sourcePath);

        // Configure custom font settings.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontDir, false);
        doc.FontSettings = fontSettings;

        // Render the document to a multipage TIFF.
        string tiffPath = Path.Combine(outputDir, "sample.tiff");
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception("TIFF rendering failed: output file not found.");
        }

        // Indicate success.
        Console.WriteLine("Document rendered to TIFF successfully: " + tiffPath);
    }
}
