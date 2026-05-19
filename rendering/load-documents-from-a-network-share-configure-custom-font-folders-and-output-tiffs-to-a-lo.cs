using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;   // Needed for LoadOptions
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string networkShareDir = Path.Combine(baseDir, "NetworkShare");
        string localOutputDir = Path.Combine(baseDir, "TiffOutput");
        string customFontDir = Path.Combine(baseDir, "CustomFonts");

        // Ensure directories exist.
        Directory.CreateDirectory(networkShareDir);
        Directory.CreateDirectory(localOutputDir);
        Directory.CreateDirectory(customFontDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        // Use a font name that may not be installed to demonstrate custom font folder usage.
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Save the document to the simulated network share location.
        string networkDocPath = Path.Combine(networkShareDir, "SampleDocument.docx");
        sourceDoc.Save(networkDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the document from the network share.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        // Configure custom font folder before loading (or after loading before rendering).
        FontSettings fontSettings = new FontSettings();
        // The folder may be empty; this demonstrates setting the folder.
        fontSettings.SetFontsFolder(customFontDir, recursive: true);
        loadOptions.FontSettings = fontSettings;

        Document loadedDoc = new Document(networkDocPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Render each page of the document to a separate TIFF image.
        // -----------------------------------------------------------------
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Example resolution; adjust as needed.
            Resolution = 300
        };

        for (int pageIndex = 0; pageIndex < loadedDoc.PageCount; pageIndex++)
        {
            // Render only the current page.
            tiffOptions.PageSet = new PageSet(pageIndex);

            string tiffPath = Path.Combine(localOutputDir, $"Page_{pageIndex + 1}.tiff");
            loadedDoc.Save(tiffPath, tiffOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // -----------------------------------------------------------------
        // 4. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine("Document rendered to TIFF images successfully.");
    }
}
