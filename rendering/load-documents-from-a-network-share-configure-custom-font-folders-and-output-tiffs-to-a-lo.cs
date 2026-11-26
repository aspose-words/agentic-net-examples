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
        // Prepare temporary directories to simulate a network share, custom fonts, and output location.
        string baseTemp = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
        string networkShareDir = Path.Combine(baseTemp, "NetworkShare");
        string customFontDir = Path.Combine(baseTemp, "CustomFonts");
        string outputDir = Path.Combine(baseTemp, "Output");

        Directory.CreateDirectory(networkShareDir);
        Directory.CreateDirectory(customFontDir);
        Directory.CreateDirectory(outputDir);

        // Create a simple source document locally.
        Document sourceDoc = new Document();
        sourceDoc.FirstSection.Body.FirstParagraph.AppendChild(new Run(sourceDoc, "Hello from Aspose.Words!"));
        string sourceDocPath = Path.Combine(networkShareDir, "Sample.docx");
        sourceDoc.Save(sourceDocPath);

        // Copy a TrueType font to the custom fonts folder (if a system font is available).
        string systemFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        string sampleFontName = "arial.ttf"; // Common Windows font.
        string systemFontPath = Path.Combine(systemFonts, sampleFontName);
        if (File.Exists(systemFontPath))
        {
            File.Copy(systemFontPath, Path.Combine(customFontDir, sampleFontName), true);
        }

        // Configure FontSettings to use the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontDir, false);

        // Load the document from the simulated network share using the configured FontSettings.
        LoadOptions loadOptions = new LoadOptions { FontSettings = fontSettings };
        Document loadedDoc = new Document(sourceDocPath, loadOptions);

        // Ensure the loaded document uses the same FontSettings when rendering.
        loadedDoc.FontSettings = fontSettings;

        // Prepare TIFF save options to render all pages.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PageSet = PageSet.All
        };

        // Save the rendered TIFF to the local output directory.
        string tiffPath = Path.Combine(outputDir, "RenderedOutput.tiff");
        loadedDoc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception("TIFF rendering failed: output file not found.");
        }

        // Optional: clean up temporary files (comment out if inspection is needed).
        // Directory.Delete(baseTemp, true);
    }
}
