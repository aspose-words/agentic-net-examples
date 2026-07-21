using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base temporary directory for the demo
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(baseDir);

        // Simulated network share folder
        string networkSharePath = Path.Combine(baseDir, "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Custom fonts folder
        string customFontFolder = Path.Combine(baseDir, "CustomFonts");
        Directory.CreateDirectory(customFontFolder);

        // Create a sample DOCX in the simulated network share
        string sourceDocPath = Path.Combine(networkSharePath, "Sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello, Aspose.Words rendering to TIFF!");
        sampleDoc.Save(sourceDocPath);

        // Configure FontSettings to use the custom fonts folder
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontFolder, false);

        // Load the document from the network share
        Document doc = new Document(sourceDocPath);
        // Apply the custom FontSettings to the loaded document
        doc.FontSettings = fontSettings;

        // Prepare output directory for TIFF files
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "Sample.tiff");

        // Render the document to a multipage TIFF (all pages are rendered by default)
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF rendering failed: file not found.");

        // Indicate success (non-interactive)
        Console.WriteLine($"Document rendered to TIFF successfully: {tiffPath}");
    }
}
