using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class EmbedFontsExample
{
    static void Main()
    {
        // Load an existing DOCX document (replace with your own path).
        Document doc = new Document("Input.docx");

        // Access the collection of fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType fonts.
        fontInfos.EmbedTrueTypeFonts = true;

        // Optionally embed system fonts as well (requires EmbedTrueTypeFonts to be true).
        fontInfos.EmbedSystemFonts = true;

        // Choose whether to embed the full font or only a subset.
        // Setting to false will embed only the glyphs used in the document (smaller file).
        // Setting to true will embed the entire font (larger file, full editability).
        fontInfos.SaveSubsetFonts = false;

        // Save the document with the embedded fonts.
        doc.Save("Output_EmbeddedFonts.docx");
    }
}
