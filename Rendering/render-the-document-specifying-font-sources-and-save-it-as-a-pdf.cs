using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to a folder that contains the custom fonts required by the document.
        string fontsDir = @"C:\MyFonts";

        // Folder where the resulting PDF will be saved.
        string artifactsDir = @"C:\Artifacts\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(artifactsDir);

        // Load or create the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example content using two different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");
        builder.Font.Name = "Arvo";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Preserve the original font sources.
        FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

        // Add a folder font source that points to the custom fonts directory.
        FolderFontSource folderFontSource = new FolderFontSource(fontsDir, true);
        FontSettings.DefaultInstance.SetFontsSources(new[] { originalFontSources[0], folderFontSource });

        // Verify that the required fonts are now available (optional).
        FontSourceBase[] currentFontSources = FontSettings.DefaultInstance.GetFontsSources();
        bool arialAvailable = currentFontSources.Any(src => src.GetAvailableFonts().Any(f => f.FullFontName == "Arial"));
        bool arvoAvailable = currentFontSources.Any(src => src.GetAvailableFonts().Any(f => f.FullFontName == "Arvo"));
        if (!arialAvailable || !arvoAvailable)
            throw new InvalidOperationException("Required fonts are not available.");

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set to true to embed full fonts, false to enable subsetting.
            EmbedFullFonts = true
        };

        // Save the document as PDF using the configured options.
        string outputPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");
        doc.Save(outputPath, pdfOptions);

        // Restore the original font sources.
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
    }
}
