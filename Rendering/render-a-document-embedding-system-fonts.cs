using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Access the collection of fonts used in the document.
        FontInfoCollection fontInfos = doc.FontInfos;

        // Enable embedding of TrueType fonts.
        fontInfos.EmbedTrueTypeFonts = true;

        // Enable embedding of system fonts (e.g., fonts installed on the OS).
        fontInfos.EmbedSystemFonts = true;

        // Optionally, embed the full fonts without subsetting.
        // Set to true to embed only the glyphs used in the document.
        fontInfos.SaveSubsetFonts = false;

        // Save the document; the fonts will be embedded according to the settings above.
        doc.Save("Output_Embedded.docx");
    }
}
