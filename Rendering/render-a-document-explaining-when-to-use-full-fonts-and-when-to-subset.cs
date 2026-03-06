using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new document and add explanatory text about font embedding.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Font Embedding in PDF");
        builder.Writeln();
        builder.Writeln("When to embed full fonts:");
        builder.Writeln("- You need to edit the PDF later and require all glyphs.");
        builder.Writeln("- The document uses a custom font that may not be available on target machines.");
        builder.Writeln();
        builder.Writeln("When to use subsetting:");
        builder.Writeln("- You want to keep the PDF file size small.");
        builder.Writeln("- Only a small subset of glyphs from the font is used.");
        builder.Writeln("- The font is large (several megabytes).");

        // Ensure the fonts used in the document are available to Aspose.Words.
        // Assume a folder named "Fonts" exists in the current directory containing any custom fonts.
        string fontsDir = Path.Combine(Environment.CurrentDirectory, "Fonts");
        FontSourceBase[] originalSources = FontSettings.DefaultInstance.GetFontsSources();
        FolderFontSource customSource = new FolderFontSource(fontsDir, true);
        FontSettings.DefaultInstance.SetFontsSources(new[] { originalSources[0], customSource });

        // Save the document to PDF with subsetting (default behavior).
        PdfSaveOptions subsetOptions = new PdfSaveOptions();
        subsetOptions.EmbedFullFonts = false; // Subset fonts to reduce file size.
        doc.Save("PdfWithSubsetFonts.pdf", subsetOptions);

        // Save the same document to PDF with full fonts embedded.
        PdfSaveOptions fullOptions = new PdfSaveOptions();
        fullOptions.EmbedFullFonts = true; // Embed the complete font files.
        doc.Save("PdfWithFullFonts.pdf", fullOptions);

        // Restore the original font sources to avoid side effects.
        FontSettings.DefaultInstance.SetFontsSources(originalSources);
    }
}
