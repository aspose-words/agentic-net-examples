using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class FontEmbeddingDemo
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add explanatory text.
        builder.Writeln("When to embed full fonts vs. subset fonts:");
        builder.Writeln();
        builder.Writeln("• Embed full fonts (EmbedFullFonts = true) when you need to edit the resulting PDF later,");
        builder.Writeln("  because all glyphs are available. This increases the file size.");
        builder.Writeln("• Use font subsetting (EmbedFullFonts = false) to keep the PDF smaller,");
        builder.Writeln("  as only the glyphs used in the document are embedded.");

        // Ensure the document uses a custom font so that embedding can be demonstrated.
        builder.Font.Name = "Arial";
        builder.Writeln("Sample text using Arial.");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("Sample text using Times New Roman.");

        // Add a custom font folder to the font sources (replace with your actual fonts folder).
        string fontsFolder = @"C:\MyFonts"; // <-- adjust path as needed
        FontSourceBase[] originalSources = FontSettings.DefaultInstance.GetFontsSources();
        FolderFontSource customSource = new FolderFontSource(fontsFolder, true);
        FontSettings.DefaultInstance.SetFontsSources(new[] { originalSources[0], customSource });

        // Save PDF with full fonts embedded.
        PdfSaveOptions fullFontOptions = new PdfSaveOptions();
        fullFontOptions.EmbedFullFonts = true; // embed every glyph
        doc.Save("Document_With_Full_Fonts.pdf", fullFontOptions);

        // Save PDF with font subsetting (default behavior).
        PdfSaveOptions subsetFontOptions = new PdfSaveOptions();
        subsetFontOptions.EmbedFullFonts = false; // embed only used glyphs
        doc.Save("Document_With_Subset_Fonts.pdf", subsetFontOptions);

        // Restore original font sources.
        FontSettings.DefaultInstance.SetFontsSources(originalSources);
    }
}
