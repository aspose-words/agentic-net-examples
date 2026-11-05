using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Simulate a network share folder.
        string networkShareFolder = Path.Combine(Directory.GetCurrentDirectory(), "NetworkShare");
        Directory.CreateDirectory(networkShareFolder);

        // Path for the source document on the simulated network share.
        string sourceDocPath = Path.Combine(networkShareFolder, "Sample.docx");

        // Create a simple source document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello from the simulated network share!");
        sourceDoc.Save(sourceDocPath);

        // Configure a custom fonts folder (empty for this example).
        string customFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontsFolder);
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontsFolder, false);

        // Load the document from the network share using the custom FontSettings.
        LoadOptions loadOptions = new LoadOptions { FontSettings = fontSettings };
        Document loadedDoc = new Document(sourceDocPath, loadOptions);

        // Prepare the output directory for TIFF files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiff");
        Directory.CreateDirectory(outputFolder);
        string outputTiffPath = Path.Combine(outputFolder, "Sample.tiff");

        // Render the document to a multipage TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        loadedDoc.Save(outputTiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputTiffPath))
        {
            throw new Exception("Failed to create the TIFF output file.");
        }

        // Indicate success.
        Console.WriteLine("Document rendered to TIFF successfully: " + outputTiffPath);
    }
}
