using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontEmbeddingExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Access the collection of fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType fonts.
        fontInfos.EmbedTrueTypeFonts = true;

        // Also embed system fonts (e.g., East Asian fonts) if needed.
        fontInfos.EmbedSystemFonts = true;

        // Save the full font data (no subsetting) – set to false to embed only used glyphs.
        fontInfos.SaveSubsetFonts = false;

        // OPTIONAL: If the document uses PostScript fonts and you want to embed them,
        // create a SaveOptions instance and enable the corresponding flag.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
        saveOptions.AllowEmbeddingPostScriptFonts = true;

        // Save the document with the updated font embedding settings.
        doc.Save("OutputDocument_WithEmbeddedFonts.docx", saveOptions);
    }
}
