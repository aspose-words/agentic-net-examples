using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Get the collection that holds information about the fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of all TrueType fonts when the document is saved.
        fontInfos.EmbedTrueTypeFonts = true;

        // Enable embedding of system fonts (useful for East Asian fonts).
        // This property only takes effect when EmbedTrueTypeFonts is true.
        fontInfos.EmbedSystemFonts = true;

        // Optionally, save the full fonts instead of subsets.
        // Set to false to embed only the glyphs used in the document.
        fontInfos.SaveSubsetFonts = false;

        // Save the document; the fonts will be embedded according to the settings above.
        doc.Save("Output_EmbeddedFonts.docx");
    }
}
