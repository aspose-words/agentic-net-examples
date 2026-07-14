using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define a temporary folder that will act as a network share.
        string networkSharePath = Path.Combine(Path.GetTempPath(), "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Define a local folder for the rendered TIFF output.
        string localOutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(localOutputPath);

        // Define a custom font folder (can be empty; Aspose.Words will fallback to system fonts).
        string customFontsPath = Path.Combine(Path.GetTempPath(), "CustomFonts");
        Directory.CreateDirectory(customFontsPath);

        // -----------------------------------------------------------------
        // 1. Create a sample document locally and save it to the "network share".
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("First page - loaded from a network share.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page - also from the network share.");

        string sourceDocPath = Path.Combine(networkSharePath, "SampleDocument.docx");
        sampleDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 2. Load the document from the network share.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);

        // -----------------------------------------------------------------
        // 3. Configure custom font settings to point to the custom font folder.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontsPath, recursive: true);
        loadedDoc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 4. Render the loaded document to a multipage TIFF image.
        // -----------------------------------------------------------------
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render at 300 DPI for reasonable quality.
            Resolution = 300
        };

        string outputTiffPath = Path.Combine(localOutputPath, "RenderedDocument.tiff");
        loadedDoc.Save(outputTiffPath, tiffOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the TIFF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputTiffPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optional: output basic information.
        Console.WriteLine($"Document loaded from: {sourceDocPath}");
        Console.WriteLine($"Custom fonts folder: {customFontsPath}");
        Console.WriteLine($"TIFF rendered to: {outputTiffPath}");
        Console.WriteLine($"Document page count: {loadedDoc.PageCount}");
    }
}
