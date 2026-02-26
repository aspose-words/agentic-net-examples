using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class RenderDocumentToPdf
{
    static void Main()
    {
        // Paths for output and custom fonts.
        string artifactsDir = @"C:\Artifacts\";
        string fontsDir = @"C:\CustomFonts\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and add text with different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");
        builder.Font.Name = "Arvo";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Preserve the original font sources.
        FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

        // Add a folder font source that points to the custom fonts folder.
        FolderFontSource folderSource = new FolderFontSource(fontsDir, true);
        FontSettings.DefaultInstance.SetFontsSources(new[] { originalFontSources[0], folderSource });

        // (Optional) Verify that the required fonts are now available.
        FontSourceBase[] currentSources = FontSettings.DefaultInstance.GetFontsSources();
        bool arialAvailable = currentSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial");
        bool arvoAvailable = currentSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Arvo");

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Embed full fonts to keep all glyphs (set to false for subsetting).
        pdfOptions.EmbedFullFonts = true;

        // Save the document as PDF using the specified options.
        string pdfPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Restore the original font sources.
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
    }
}
