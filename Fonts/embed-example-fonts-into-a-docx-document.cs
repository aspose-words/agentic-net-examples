using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class EmbedFontsExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Access the collection of fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType and system fonts.
        // Also save only the subset of each font that is actually used.
        bool embedAll = true;
        fontInfos.EmbedTrueTypeFonts = embedAll;   // Embed TrueType fonts.
        fontInfos.EmbedSystemFonts = embedAll;    // Embed system fonts (e.g., Arial, Times New Roman).
        fontInfos.SaveSubsetFonts = embedAll;     // Save only the used glyphs.

        // Save the document with the embedded fonts.
        doc.Save("OutputDocument.docx");
    }
}
