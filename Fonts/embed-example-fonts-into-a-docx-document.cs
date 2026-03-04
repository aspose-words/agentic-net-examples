using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class EmbedFontsExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Access the collection of fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType fonts.
        fontInfos.EmbedTrueTypeFonts = true;

        // Also embed system fonts (required when EmbedTrueTypeFonts is true).
        fontInfos.EmbedSystemFonts = true;

        // Optionally, embed the full font files instead of subsets.
        // Set to true to embed only the glyphs used in the document (smaller file).
        // Set to false to embed the entire font (larger file, full editability).
        fontInfos.SaveSubsetFonts = false;

        // Save the document with the embedded fonts.
        doc.Save("Output_EmbeddedFonts.docx");
    }
}
