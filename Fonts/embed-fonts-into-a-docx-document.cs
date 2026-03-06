using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Get the collection that holds information about the fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of TrueType fonts when the document is saved.
        fontInfos.EmbedTrueTypeFonts = true;

        // Optionally embed system fonts as well (useful for East Asian languages).
        fontInfos.EmbedSystemFonts = true;

        // Save the full font files (set to false to embed only subsets).
        fontInfos.SaveSubsetFonts = false;

        // Save the document; the fonts will be embedded according to the settings above.
        doc.Save("Output_EmbeddedFonts.docx");
    }
}
