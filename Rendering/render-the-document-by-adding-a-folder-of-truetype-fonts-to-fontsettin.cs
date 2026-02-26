using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontEmbeddingExample
{
    static void Main()
    {
        // Path to a folder that contains TrueType fonts.
        string fontsDir = @"C:\MyFonts";

        // Path where the resulting PDF will be saved.
        string outputPdf = @"C:\Output\Document.pdf";

        // Choose whether to embed the full font (true) or only the subset used in the document (false).
        bool embedFullFonts = true;

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text using a custom font that resides in the fonts folder.
        builder.Font.Name = "Amethysta"; // Example custom font.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Preserve the original font sources so we can restore them later.
        FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

        // Create a folder font source that points to the custom fonts directory.
        // The second argument (true) enables recursive search in subfolders.
        FolderFontSource folderFontSource = new FolderFontSource(fontsDir, true);

        // Combine the original sources with the new folder source.
        FontSourceBase[] updatedFontSources = originalFontSources.Concat(new[] { folderFontSource }).ToArray();

        // Apply the updated font sources to the default FontSettings instance.
        FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

        // Verify that the custom font is now available (optional, can be removed in production).
        bool fontAvailable = FontSettings.DefaultInstance.GetFontsSources()
                               .SelectMany(src => src.GetAvailableFonts())
                               .Any(f => f.FullFontName.Equals("Amethysta", StringComparison.OrdinalIgnoreCase));

        if (!fontAvailable)
            throw new InvalidOperationException("The required font was not found in the specified font sources.");

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When true, the entire font file is embedded; when false, only the used glyphs are embedded.
            EmbedFullFonts = embedFullFonts
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPdf, pdfOptions);

        // Restore the original font sources to avoid side effects on other operations.
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
    }
}
