using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontEmbeddingExample
{
    static void Main()
    {
        // Path to the folder that contains the TrueType fonts to be used.
        string fontsDir = @"C:\MyFonts";

        // Path where the resulting PDF will be saved.
        string outputPdf = @"C:\Output\Result.pdf";

        // Choose whether to embed the full font (true) or only the subset used (false).
        bool embedFullFonts = true;

        // Create a new empty document.
        Document doc = new Document();

        // Build some sample content with a custom font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFont"; // Assume this font exists in the fonts folder.
        builder.Writeln("Hello world with a custom TrueType font!");

        // Preserve the original font sources so we can restore them later.
        FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

        // Create a font source that points to the folder with our TrueType fonts.
        FolderFontSource folderFontSource = new FolderFontSource(fontsDir, true);

        // Add the new font source to the existing sources.
        FontSourceBase[] updatedFontSources = new FontSourceBase[originalFontSources.Length + 1];
        originalFontSources.CopyTo(updatedFontSources, 0);
        updatedFontSources[originalFontSources.Length] = folderFontSource;
        FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

        // Verify that the custom font is now available (optional).
        // bool fontAvailable = FontSettings.DefaultInstance.GetFontsSources()
        //     .Any(src => src.GetAvailableFonts().Any(f => f.FullFontName == "CustomFont"));

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When true, embed every glyph of each font; when false, embed only the subset used.
            EmbedFullFonts = embedFullFonts,

            // Optional: control which fonts are embedded.
            // FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPdf, pdfOptions);

        // Restore the original font sources to avoid side effects on other operations.
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
    }
}
